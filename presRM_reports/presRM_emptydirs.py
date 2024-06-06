import secret
import pyPreservica

entity = pyPreservica.EntityAPI(username=secret.username,password=secret.password,server="unilever.preservica.com")

nlref = "5f76e749-ddf3-4d3d-a31c-bd07864e9dc2"
ukref = "bf614a5d-10db-4ff8-addb-2c913b3149eb"

olduktitle = "FOLDER EMPTY; ALL BOXES SIGNED OFF --- "
newuktitle = "EMPTY --- "
newukdesc = "Folder empty. All Boxes Signed Off --- "
newnltitle = "EMPTY --- "
oldnltitle = "FOLDER EMPTY; DO NOT USE --- "
newnldesc = "Folder Empty, do note use; All archive material ingested into QI from Project Snakeboot 2023 --- "
oldnldesc = "All ARCHIVE material ingested into QI from Project Snakeboot 2023 --- "

def empty_iter(ref):
    ents = list(entity.descendants(ref))
    if len(ents) == 0:
        ent = entity.folder(ref)
        # if oldnltitle in ent.title:
        #     ent.title = ent.title.replace(oldnltitle,newnltitle)
        # if ent.description:
        #     if oldnldesc in ent.description: ent.title.replace(oldnldesc,newnldesc)
        #     else: ent.description = newnldesc + ent.description
        # else: ent.description = newnldesc
        # ent = entity.save(ent)
        print(ent.title)
    else:
        for e in ents:            
            if str(e.entity_type) == "EntityType.FOLDER":
                empty_iter(e.reference)

empty_iter(nlref)