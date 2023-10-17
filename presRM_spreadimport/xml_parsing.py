from datetime import datetime
from lxml import etree

def text_formatting(value):
    value = str(value)
    if value == "False" or value == "NaN" or value == "nan" or value == "NaT" or value == "": value = ""
    else: 
        value = str(value).strip()
        value = value.replace("&","&amp;")
        value = value.replace(">","&gt;")
        value = value.replace("<","&lt;")
        value = value.replace("'","&apos;")
        value = value.replace("\"","&quot;")
    return value

def date_formatting(vardate):
    vardate = str(vardate)
    if vardate == "False" or vardate == "NaN" or vardate == "nan" or vardate == "NaT" or vardate == "": vardate = datetime.now().strftime("%Y-%m-%dT%H:%M:%S.000Z")
    else: vardate = datetime.strptime(vardate,"%Y-%m-%d %H:%M:%S").strftime("%Y-%m-%dT%H:%M:%S.000Z")
    return vardate

class XML_Parse():
    def __init__(self,NS,box_dict=None,item_dict=None,tm_dict=None):
        self.NS = NS
        self.box_dict = box_dict
        self.item_dict = item_dict
        self.tm_dict = tm_dict

    def metadata_parse_TM(self,record):
        ##Customise to remove all special characters...
        Location = text_formatting(record.get('Location'))
        Box = text_formatting(record.get('Box'))
        Reference = text_formatting(record.get('Box'))
        Country = text_formatting(record.get('Country'))
        Trademark = text_formatting(record.get('Trademark Name'))
        Registration = text_formatting(record.get('Registration Number'))
        Type = text_formatting(record.get('Type of Document'))
        Parties = text_formatting(record.get('Parties Involved'))
        Add = text_formatting(record.get('Additional Information'))
        xml = f'<tm:tm xmlns:tm="{self.NS}" xmlns="{self.NS}">' \
            f"<tm:location>{Location}</tm:location>" \
            f"<tm:box>{Box}</tm:box>" \
            f"<tm:reference>{Reference}</tm:reference>" \
            f"<tm:country>{Country}</tm:country>" \
            f"<tm:trademark>{Trademark}</tm:trademark>" \
            f"<tm:registration>{Registration}</tm:registration>" \
            f"<tm:type>{Type}</tm:type>" \
            f"<tm:parties>{Parties}</tm:parties>" \
            f"<tm:add>{Add}</tm:add>" \
            "</tm:tm>"    
        try: xml_test = etree.fromstring(xml)
        except Exception as e: 
            print(e)
            print(xml)
        
        return xml

    def metadata_parse(self,box_dict,item_dict=None,transfer_dict=None,type_flag=None,up_flag=None):
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
        client = ""
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
        
        labtype = ""
        labnotebook = ""
        author = ""
        department = ""
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
        alternativereference = ""

        if type_flag == "Box":
            objtype = "Box"
            title = text_formatting(box_dict.get('Box Reference'))
            desc = text_formatting(box_dict.get('Description'))        
            loc = text_formatting(box_dict.get('Location'))
            weight = text_formatting(box_dict.get('Weight'))
            altcont = text_formatting(box_dict.get('Alternative Contact'))
            client = text_formatting(box_dict.get('Client'))
            if up_flag != "Legal": statusdate = date_formatting(box_dict.get('Status Date'))
            else: statusdate = ""
            boxtype = text_formatting(box_dict.get('Box Type'))
            area = text_formatting(box_dict.get('Area'))
            defaultloc = text_formatting(box_dict.get('Home Location'))
            coverdates = text_formatting(box_dict.get('Covering Dates'))
            notes = text_formatting(box_dict.get('Notes'))

            #Transfer_Dict only Get's fed into Box Level
            #If Statement to filter out Legal.
            if up_flag != "Legal":
                transdate = date_formatting(transfer_dict.get('Transfer Date'))
                transaccpt = text_formatting(transfer_dict.get('Accepted By'))
                transby = text_formatting(transfer_dict.get('Transferred By'))
                transcharge = text_formatting(transfer_dict.get('Chargeable'))
                transnotes = text_formatting(transfer_dict.get('Transfer Ntoes'))
            else:
                transdate = ""
                transaccpt = ""
                transby = ""
                transcharge = ""
                transnotes = ""
            
        elif type_flag == "Item":
            objtype = "Item"
            #Box Information Recurses down to Item Level
            loc = text_formatting(box_dict.get('Location'))
            altcont = text_formatting(box_dict.get('Alternative Contact'))
            client = text_formatting(box_dict.get('Client'))
            statusdate = date_formatting(box_dict.get('Status Date'))
            area = text_formatting(box_dict.get('Area'))
            defaultloc = text_formatting(box_dict.get('Home Location'))
            #Item Level Info
            title = text_formatting(item_dict.get('Item Reference'))
            desc = text_formatting(item_dict.get('Description'))        
            itemtype = text_formatting(item_dict.get('Item Type'))
            numberofitems = text_formatting(item_dict.get('Number of Items'))
            coverdates = text_formatting(item_dict.get('Covering Dates'))
            format = text_formatting(item_dict.get('Format'))
            notes = text_formatting(item_dict.get('Notes'))

        elif type_flag == "Lab":
            objtype = "Item"
            #Box Information Recurses down to Item Level
            loc = text_formatting(box_dict.get('Location'))
            altcont = text_formatting(box_dict.get('Alternative Contact'))
            client = text_formatting(box_dict.get('Client'))
            statusdate = date_formatting(box_dict.get('Status Date'))
            area = text_formatting(box_dict.get('Area'))
            defaultloc = text_formatting(box_dict.get('Home Location'))
            #Lab - Base
            title = text_formatting(box_dict.get('Title'))
            desc = text_formatting(box_dict.get('Description'))        
            coverdates = text_formatting(item_dict.get('Covering Dates'))
            itemtype = text_formatting(item_dict.get('Item Type'))
            notes = text_formatting(item_dict.get('Notes'))
            #Lab - Unq
            labtype = text_formatting(item_dict.get('Lab Type'))
            labnotebook = text_formatting(item_dict.get('Lab Notebook Number'))
            author = text_formatting(item_dict.get('Author'))
            department = text_formatting(item_dict.get('Department'))
            site = text_formatting(item_dict.get('Site'))
            dateofissue = date_formatting(item_dict.get('Date of Issue'))
            dateofcompletion = date_formatting(item_dict.get('Date of Completion'))
        elif type_flag == "Legal":
            objtype = "Item"
            #Box Information Recurses down to Item Level
            loc = text_formatting(box_dict.get('Location'))
            altcont = text_formatting(box_dict.get('Alternative Contact'))
            client = text_formatting(box_dict.get('Client'))
            #statusdate = date_formatting(box_dict.get('Status Date'))
            area = text_formatting(box_dict.get('Area'))
            defaultloc = text_formatting(box_dict.get('Home Location'))
            #Legal - Base
            itemtype = text_formatting(item_dict.get('Item Type'))
            title = text_formatting(item_dict.get('Agreement Reference'))
            desc = text_formatting(item_dict.get('Description'))        
            coverdates = text_formatting(item_dict.get('Covering Dates'))
            notes = text_formatting(item_dict.get('Notes'))
            #Legal - Unq
            legaldept = text_formatting(transfer_dict.get('Top-Level')) # Makes Special Call on Transfer_Dict
            legaldoc = text_formatting(item_dict.get('Agreement Type'))
            legaldocother = text_formatting(item_dict.get('If Other, please state'))
            copy = text_formatting(item_dict.get('Copy'))
            party1 = text_formatting(item_dict.get('Party Name 1'))
            party2 = text_formatting(item_dict.get('Party Name 2'))
            party3 = text_formatting(item_dict.get('Party Name 3'))
            party4 = text_formatting(item_dict.get('Party Name 4'))
            dateofagreement = date_formatting(item_dict.get('Date of Agreement'))
            dateoftermination = date_formatting(item_dict.get('Date of Termination'))
            amendments = text_formatting(item_dict.get('Amendments'))
            value = text_formatting(item_dict.get('Value'))
            currency = text_formatting(item_dict.get('Currency'))
            if legaldept == "UKI":
                datesubmitted = date_formatting(item_dict.get('Date Submitted (UKI)'))
                agreementobjectives = text_formatting(item_dict.get('Agreement Objectives (UKI)'))
            else: datesubmitted = ""; agreementobjectives = ""

        # elif type_flag == "USCC": ### USCC Needs Template Defining...
        #     #USCC Info
        #     itemtype = "USCC"
        #     usccagreementtype = text_formatting(row['uscctype'])
        #     agreementowner = text_formatting(row['usccowner'])
        #     destructiondate = date_formatting(row['usccdestruct'])
        #     agreementend = date_formatting(row['usccend'])
        #     docnum = text_formatting(row['usccref'])
        #     agreementstart = date_formatting(row['usccstart'])
        #     vendordesc = text_formatting(row['usccvenddesc'])
        #     vendorcode = text_formatting(row['usccvendname'])
        #     upccode = text_formatting(row['usccuuc'])
        #     usccdatesubmitted = text_formatting(row['datesubmitted'])

        elif type_flag == "TM":
            objtype = "Item"
        else: print('Invalid Type'); pass

        xml_parse = f'<rm:rm xmlns:rm="{self.NS}" xmlns="{self.NS}">'\
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
                    f"<rm:agreementtype>{labtype}</rm:agreementtype>" \
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
                    f"<rm:alternativereference>{alternativereference}</rm:alternativereference>" \
                    "</rm:legacyinfo>" \
                    "</rm:rm>"
        
        try: xml_test = etree.fromstring(xml_parse)
        except Exception as e:
            print(e)

        return xml_parse,loc