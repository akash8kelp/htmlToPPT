import zipfile, shutil, os
from lxml import etree
from pathlib import Path

def unzip_pptx(path, out_dir):
    with zipfile.ZipFile(path, 'r') as zf:
        zf.extractall(out_dir)

def zip_dir(folder, out_path):
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(folder):
            for f in files:
                fp = os.path.join(root, f)
                zf.write(fp, os.path.relpath(fp, folder))

def merge_pptx_xml(base_pptx, others, output):
    base_dir = "tmp_base"
    shutil.rmtree(base_dir, ignore_errors=True)
    unzip_pptx(base_pptx, base_dir)

    pres_xml = Path(base_dir, "ppt/presentation.xml")
    tree = etree.parse(str(pres_xml))
    ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main', 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

    # find existing max slideId
    sldIdLst = tree.xpath("//p:sldIdLst", namespaces=ns)[0]
    last_id = max(int(s.get("id")) for s in sldIdLst.xpath("p:sldId", namespaces=ns))
    slide_count = len(sldIdLst.xpath("p:sldId", namespaces=ns))
    
    # Also need to handle relationships
    pres_rels_xml = Path(base_dir, "ppt/_rels/presentation.xml.rels")
    rels_tree = etree.parse(str(pres_rels_xml))
    last_rid = max(int(r.get("Id").replace("rId", "")) for r in rels_tree.getroot())


    for idx, f in enumerate(others, start=1):
        dirN = f"tmp_{idx}"
        shutil.rmtree(dirN, ignore_errors=True)
        unzip_pptx(f, dirN)

        slide_files = sorted(Path(dirN, "ppt/slides").glob("slide*.xml"))
        for s in slide_files:
            slide_count += 1
            new_name = f"slide{slide_count}.xml"
            shutil.copy(s, Path(base_dir, "ppt/slides", new_name))
            
            # also copy rels and media
            rel_file = s.parent.parent / "slides/_rels" / (s.stem + ".xml.rels")
            if rel_file.exists():
                shutil.copy(rel_file, Path(base_dir, "ppt/slides/_rels", new_name + ".rels"))
            
            media_folder = s.parent.parent / "media"
            if media_folder.exists():
                for media in media_folder.glob("*"):
                    shutil.copy(media, Path(base_dir, "ppt/media"))

            # append new slide ID
            last_id += 1
            last_rid +=1
            
            new_sldId = etree.Element(
                "{http://schemas.openxmlformats.org/presentationml/2006/main}sldId",
                id=str(last_id),
            )
            new_sldId.set("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", f"rId{last_rid}")
            sldIdLst.append(new_sldId)
            
            # Add relationship for the new slide
            new_rel = etree.Element("Relationship", Id=f"rId{last_rid}", Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide", Target=f"slides/{new_name}")
            rels_tree.getroot().append(new_rel)


    tree.write(str(pres_xml), xml_declaration=True, encoding="utf-8", standalone="yes")
    rels_tree.write(str(pres_rels_xml), xml_declaration=True, encoding="utf-8", standalone="yes")

    zip_dir(base_dir, output)
    print(f"✅ merged -> {output}")
    # Clean up temp directories
    shutil.rmtree(base_dir, ignore_errors=True)
    for idx in range(1, len(others) + 1):
        shutil.rmtree(f"tmp_{idx}", ignore_errors=True)


if __name__ == '__main__':
    # Define the folder and output file
    source_folder = "sampleHTML"
    output_file = "merged_all_slides.pptx"

    # Find all .pptx files in the source folder, excluding temporary files
    pptx_files = sorted([
        os.path.join(source_folder, f)
        for f in os.listdir(source_folder)
        if f.endswith('.pptx') and not f.startswith('~$')
    ])

    if len(pptx_files) < 2:
        print(f"Error: Found fewer than 2 PPTX files in '{source_folder}'. Nothing to merge.")
    else:
        # Use the first file as the base and the rest for merging
        base_file = pptx_files[0]
        files_to_merge = pptx_files[1:]
        
        print(f"Base presentation: {base_file}")
        print("Presentations to merge:")
        for f in files_to_merge:
            print(f"  - {f}")
        
        merge_pptx_xml(
            base_file,
            files_to_merge,
            output_file
        )
