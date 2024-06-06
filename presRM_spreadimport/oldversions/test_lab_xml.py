from lxml import etree
from pyPreservica import *

def xml_parsing():
    transdate = ""
    transaccpt = ""
    transby = ""
    transcharge = ""
    transnotes = ""

    objtype = ""
    title = ""
    desc = ""

    loc = ""
    weight = ""
    altcont = ""
    client = "Test"
    statusdate = ""
    boxtype = ""
    area = ""
    defaultloc = ""
    coverdates = ""
    notes = ""

    itemtype = ""
    numberofitems = ""
    coverdates = ""
    format = ""

    labagreementtype = ""
    labnotebook = ""
    author = "TEst"
    department = "Test"
    site = ""
    dateofissue = ""
    dateofcompletion  = ""

    usccagreementtype = ""
    legaldept = ""
    legaldoc = ""
    legaldocother = ""
    copy = ""
    party1 = ""
    party2 = ""
    party3 = ""
    party4 = ""
    dateofagreement = ""
    dateoftermination  = ""
    amendments = ""
    value = ""
    currency = ""
    datesubmitted = ""
    agreementobjectives = ""
            
    usccagreementtype = ""
    agreementowner = ""
    destructiondate = ""
    agreementend = ""
    docnum = ""
    agreementstart = ""
    vendordesc = ""
    vendorcode = ""
    upccode = ""
    usccdatesubmitted = ""

    #Legacy Info - Will always be set to Null.

    legacycreatedby = ""
    legacycreateddate = ""
    legacyid = ""
    legacypid = ""
    altref = ""

    xml_parse = f'<rm:rm xmlns:rm="{RMNS}" xmlns="{RMNS}">'\
                        "<rm:recordinfo>" \
                        f"<rm:statusdate>{statusdate}</rm:statusdate>" \
                        f"<rm:coverdate>{coverdates}</rm:coverdate>" \
                        f"<rm:disposition>false</rm:disposition>" \
                        f"<rm:objtype>{objtype}</rm:objtype>" \
                        f"<rm:recordnotes>{notes}</rm:recordnotes>" \
                        "</rm:recordinfo>" \
                        "<rm:clientinfo>" \
                        f"<rm:client>{client}</rm:client>"\
                        f"<rm:alternatecontact>{altcont}</rm:alternatecontact>" \
                        "</rm:clientinfo>" \
                        "<rm:boxinfo>" \
                        f"<rm:weight>{weight}</rm:weight>" \
                        f"<rm:area>{area}</rm:area>" \
                        f"<rm:boxlocation>{loc}</rm:boxlocation>" \
                        f"<rm:defaultlocation>{defaultloc}</rm:defaultlocation>" \
                        f"<rm:boxtype>{boxtype}</rm:boxtype>" \
                        "</rm:boxinfo>" \
                        "<rm:iteminfo>" \
                        f"<rm:itemtype>{itemtype}</rm:itemtype>" \
                        f"<rm:format>{format}</rm:format>" \
                        f"<rm:numberofitems>{numberofitems}</rm:numberofitems>" \
                        "<rm:legalitem>" \
                        f"<rm:legaldept>{legaldept}</rm:legaldept>" \
                        f"<rm:legaldoc>{legaldoc}</rm:legaldoc>" \
                        f"<rm:legaldocother>{legaldocother}</rm:legaldocother>" \
                        f"<rm:copy>{copy}</rm:copy>" \
                        f"<rm:partyname1>{party1}</rm:partyname1>" \
                        f"<rm:partyname2>{party2}</rm:partyname2>" \
                        f"<rm:partyname3>{party3}</rm:partyname3>" \
                        f"<rm:partyname4>{party4}</rm:partyname4>" \
                        f"<rm:dateofagreement>{dateofagreement}</rm:dateofagreement>" \
                        f"<rm:dateoftermination>{dateoftermination}</rm:dateoftermination>" \
                        f"<rm:amendments>{amendments}</rm:amendments>" \
                        f"<rm:value>{value}</rm:value>" \
                        f"<rm:currency>{currency}</rm:currency>" \
                        f"<rm:datesubmitted>{datesubmitted}</rm:datesubmitted>" \
                        f"<rm:agreementobjectives>{agreementobjectives}</rm:agreementobjectives>" \
                        "</rm:legalitem>" \
                        "<rm:usccitem>" \
                        f"<rm:agreementtype>{usccagreementtype}</rm:agreementtype>" \
                        f"<rm:upccode>{upccode}</rm:upccode>" \
                        f"<rm:vendorcode>{vendorcode}</rm:vendorcode>" \
                        f"<rm:vendordesc>{vendordesc}</rm:vendordesc>" \
                        f"<rm:agreementowner>{agreementowner}</rm:agreementowner>" \
                        f"<rm:agreementstart>{agreementstart}</rm:agreementstart>" \
                        f"<rm:agreementend>{agreementend}</rm:agreementend>" \
                        f"<rm:docnumber>{docnum}</rm:docnumber>" \
                        f"<rm:destructiondate>{usccdatesubmitted}</rm:destructiondate>" \
                        f"<rm:destructiondate>{destructiondate}</rm:destructiondate>" \
                        "</rm:usccitem>" \
                        "<rm:labitem>" \
                        f"<rm:agreementtype>{labagreementtype}</rm:agreementtype>" \
                        f"<rm:notebook>{labnotebook}</rm:notebook>" \
                        f"<rm:author>{author}</rm:author>" \
                        f"<rm:department>{department}</rm:department>" \
                        f"<rm:site>{site}</rm:site>" \
                        f"<rm:dateofissue>{dateofissue}</rm:dateofissue>" \
                        f"<rm:dateofcompletion>{dateofcompletion}</rm:dateofcompletion>" \
                        "</rm:labitem>" \
                        "</rm:iteminfo>"\
                        "<rm:transferinfo>" \
                        f"<rm:transferdate>{transdate}</rm:transferdate>" \
                        f"<rm:transferaccepted>{transaccpt}</rm:transferaccepted>" \
                        f"<rm:transfernotes>{transnotes}</rm:transfernotes>" \
                        f"<rm:transfercharge>{transcharge}</rm:transfercharge>" \
                        f"<rm:transferredby>{transby}</rm:transferredby>" \
                        "</rm:transferinfo>" \
                        "<rm:legacyinfo>" \
                        f"<rm:createdby>{legacycreatedby}</rm:createdby>" \
                        f"<rm:createddate>{legacycreateddate}</rm:createddate>" \
                        f"<rm:legacyid>{legacyid}</rm:legacyid>" \
                        f"<rm:legacypid>{legacypid}</rm:legacypid>" \
                        f"<rm:alternativereference>{altref}</rm:alternativereference>" \
                        "</rm:legacyinfo>" \
                        "</rm:rm>"
    return xml_parse


username = "test@email.com"
password = "SecretPassword"

entity = EntityAPI(username={username},password={password}, \
                    tenant="UARM",server="unilever.preservica.com")

ref = "672710d7-07d2-4e7a-af5c-b85bdbc7a934"
RMNS = "http://rm.unilever.co.uk/schema"

folder = entity.folder(ref)
xml_string = entity.metadata_for_entity(folder,RMNS)

#print(xml_string)

xml_a = etree.fromstring(xml_string)
xml_string_b = xml_parsing()
xml_b = etree.fromstring(xml_string_b)


def xml_update(a,b):
    for bchild in b.findall('./'):
        achild = a.find('./' + bchild.tag)
        if achild is not None: 
            if bchild.text:
                achild.text = bchild.text
        else: print(bchild.tag)
        if bchild.getchildren(): 
            xml_update(achild,bchild)
            #print(f'could not find: {bchild.tag}')
    return a

new_xml = xml_update(xml_a, xml_b)

print(etree.tostring(new_xml))
#print(xml_b)

