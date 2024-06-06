from pyPreservica import *
import pandas as pd
from lxml import etree
import requests
test = "ad831067-85e1-4324-96f5-d0bed78eab0a"
import secret

entity = EntityAPI(username=secret.username,password=secret.password,server="unilever.preservica.com")


url = f"https://unilever.preservica.com/api/entity/information-objects/{test}"
headers={"accept":"application/xml;charset=UTF-8","Preservica-Access-Token":entity.token}

r = requests.get(url,headers=headers)
REVNS = "http://rm.unilever.co.uk/review/v1"

def xml_merge(xml_a,xml_b):
    for b_child in xml_b.findall('./'):
        a_child = xml_a.find('./' + b_child.tag)
        if a_child is not None:
            if b_child.text:
                a_child.text = b_child.text
            else: a_child.text = None
        else:
            #print(b_child.tag)
            pass
        if b_child.getchildren(): 
            xml_merge(a_child,b_child)
    return etree.tostring(xml_a)

test_string = """
<ns0:rmreview xmlns="http://preservica.com/XIP/v7.0" xmlns:ns0="http://rm.unilever.co.uk/review/v1">
<ns0:datereview>2015-01-08T00:00:00</ns0:datereview>
<ns0:reviewnumber>test</ns0:reviewnumber>
<ns0:datereviewsent></ns0:datereviewsent>
<ns0:reviewsentby>Me!</ns0:reviewsentby>
<ns0:decision></ns0:decision>
<ns0:datedecision></ns0:datedecision>
<ns0:authorisedby></ns0:authorisedby>
<ns0:formerrsi/>
<ns0:reviewnotes/>
</ns0:rmreview>"""

test_xml = etree.fromstring(test_string.encode('utf-8'))

#print(r.text)
xml_d = etree.ElementTree(etree.fromstring(r.text.encode('utf-8')))
root = xml_d.getroot()
for n,fs in enumerate(xml_d.findall(f'//Fragment[@schema="{REVNS}"]',root.nsmap)):
    if n == 0:
        mref = str(fs.text).rsplit("/",1)[-1]
        meta_url = f"https://unilever.preservica.com/api/entity/information-objects/{test}/metadata/{mref}"
        r = requests.get(meta_url,headers=headers)
        meta_xml_d = etree.ElementTree(etree.fromstring(r.text.encode('utf-8')))
        md =meta_xml_d.find("//{http://preservica.com/XIP/v7.0}Content/*")
        new_xml = xml_merge(md,test_xml)
        print(new_xml)
    else: pass

#
filepath = r"C:\Users\Chris.Prince\Downloads\RR2023 Bulk Updated Reviews Extended_Accidently overwritten (1).xlsx"

#this ended up not being necessary.
asset = entity.asset(test)
for meta in entity.all_metadata(asset):
    if meta[0] == REVNS:
        xml_string = meta[1]
        break

print(xml_string)
xml_data = etree.fromstring(xml_string)
date_review = xml_data.find('{http://rm.unilever.co.uk/review/v1}datereview')
date_review.text = "2015-01-08T00:00:00.000Z"
etree.tostring(xml_data).decode('utf-8')

real_date = "2024-03-05T00:00:00.000Z"

#this is the call to update the metadata based on the above.
#entity.update_metadata(asset,REVNS,etree.tostring(xml_data).decode('utf-8'))


"""
The above was what I did to work things out. It's always good to just simply test by trying things.

Essentially what I'm doing above is making a call on the API

I've had to do a manual call on the API here, as PyPreservica was unsuitable for the task of  

The below is the code that fixed. It takes the above and loops through the spreadsheet you've provided. In that spreadsheet I've merged the prior review metadata together.


"""

print('\n\n\nTest Code Ends Here\n\n\n')
time.sleep(5)

#This is my XML_merge function, it takes one XML and merges it with another. I've explained this elsewhere :P
#I've modified this slightly so that if b_child is empty it will override a_child with Nothing. 
#Note to self, use this version from now on.
def xml_merge(xml_a,xml_b):
    for b_child in xml_b.findall('./'):
        a_child = xml_a.find('./' + b_child.tag)
        if a_child is not None:
            if b_child.text:
                a_child.text = b_child.text
            else: a_child.text = None
        else:
            #print(b_child.tag)
            pass
        if b_child.getchildren(): 
            xml_merge(a_child,b_child)
    return etree.tostring(xml_a)

# I'm sneaky here and use pyPreservica to login into Preservica and fetch the Access Token, rather than manually call on it myself.
# Note: this is not good practice to do, as the token will only last 15 minutes, before needing to be refreshed... But I'm lazy. :)

token = entity.token

# I set the headers for the manual request on the API - these will be constant, as we're making the same calls on the same API endpoint, 
# If we were making different calls, we might have to change them in our loop.
headers={"Content-Type":"application/xml;charset=UTF-8","Preservica-Access-Token":token}

# The Filepath of the Spreadsheet you sent, with Review data merged in.
filepath = r"C:\Users\Chris.Prince\Downloads\RR2023 Bulk Updated Reviews Extended_Accidently overwritten (1).xlsx"

# Set the XML namespace as a constant, as it will be constant.
REVNS = "http://rm.unilever.co.uk/review/v1"

# I take the spreadsheet import into a Dataframe, then export it to a List of Dictionaries.
df = pd.read_excel(filepath)
records = df.to_dict('records')

# I then loop over the list of dictionaries, which will be the records we want to update.
for record in records:
    pres_ref = record.get('Reference')
    # An XML string is created using the data from the Dict 
    xml_string = f"""
            <ns0:rmreview xmlns="http://preservica.com/XIP/v7.0" xmlns:ns0="{REVNS}">
                <ns0:datereview>{record.get('reviewsdate')}</ns0:datereview>
                <ns0:reviewnumber/>
                <ns0:datereviewsent/>
                <ns0:reviewsentby>{record.get('reviewssentby')}</ns0:reviewsentby>
                <ns0:decision>{record.get('reviewsaction')}</ns0:decision>
                <ns0:datedecision/>
                <ns0:authorisedby>{record.get('reviewsauthorisedby')}</ns0:authorisedby>
                <ns0:formerrsi/>
                <ns0:reviewnotes/>
            </ns0:rmreview>"""
    #
    xml_data = etree.fromstring(xml_string.encode('utf-8'))
    #I form the URL then make a get request on the API to return the data.
    url = f"https://unilever.preservica.com/api/entity/information-objects/{pres_ref}"
    r = requests.get(url,headers=headers)
    #I load the returned data (which is a string) into an XML, note this isn't the metadata we're after yet.
    xip_data = etree.ElementTree(etree.fromstring(r.text.encode('utf-8')))
    root = xip_data.getroot()
    #This is finding the XML element "Fragment" with our XML namespace (so it only pulls through the right fragment).
    for n,fs in enumerate(xip_data.findall(f'//Fragment[@schema="{REVNS}"]',root.nsmap)):
        #I'm also using enumerate, to prevent it override all of the fragments, by simply running the script within an if n == 0, IE the first fragment.
        if n == 0:
            #I'm fetching the Metadata Reference code here
            mref = str(fs.text).rsplit("/",1)[-1]
            #I then form a new URL to call to make another call on the API, to retrieve the actual metadata.
            meta_url = f"https://unilever.preservica.com/api/entity/information-objects/{pres_ref}/metadata/{mref}"
            print(meta_url)
            meta_r = requests.get(meta_url,headers=headers)
            #Parse the returned string to an XML element
            pres_meta_xml = etree.ElementTree(etree.fromstring(meta_r.text.encode('utf-8')))
            # I quickly build a new XML document using lxml, that I will replace
            # I could have potentially reused the existing XML that gets downloaded
            xml_wrapper = etree.Element('MetadataContainer',{"schemaUri":REVNS,'xmlns':"http://preservica.com/XIP/v6.9"})
            etree.SubElement(xml_wrapper,'Ref').text = mref
            etree.SubElement(xml_wrapper,'Entity').text = pres_ref
            content_wrapper = etree.SubElement(xml_wrapper, "Content")
            # Here I'm retrieving the actual metadata section of the XML. this was important as the additional bits metadata are part of the 
            # XIP metadata not our RM data, I've added a template of how this appear.
            pres_meta_data = pres_meta_xml.find("//{http://preservica.com/XIP/v7.0}Content/*")
            #I'm calling on my XML Merge function to merge together the data I've just retrieved with the string I created before.
            new_xml = xml_merge(pres_meta_data,xml_data)
            new_tree = etree.fromstring(new_xml)
            #This then gets appended into the Content section of the new XML I created
            content_wrapper.append(new_tree)
            #I print the new wrapper for visibility,
            print(etree.tostring(xml_wrapper,encoding="utf-8"))
            print()
            #Finally I send of a PUT request to make the changes, note that data is our new wrapper.
            put_request = requests.put(meta_url,headers=headers,data=etree.tostring(xml_wrapper,encoding="utf-8").decode('utf-8'))
            #Printing the returned data, so I know if it goes through or not :)
            print(put_request.status_code)
            print(put_request.text)

        #This else is here as we're only operating on the first result. 
        else: pass

# The End!
        

###