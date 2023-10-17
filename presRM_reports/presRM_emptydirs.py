import secret
import pyPreservica

entity = pyPreservica.EntityAPI(username=secret.username,password=secret.password,server="unilever.preservica.com")

#ref = "ade77e29-697d-4d45-a97c-ed258bfb95b8"
ref = "095b294d-6b83-4684-953a-8c89b11d2956"


def empty_iter(ref):
    ents = list(entity.descendants(ref))
    if len(ents) == 0:
        print(f'Empty - {ref}')
    else:
        for e in ents:
            if str(e.entity_type) == "EntityType.FOLDER":
                empty_iter(e.reference)
empty_iter(None)