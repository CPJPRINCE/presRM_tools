from pyPreservica import *
import secret

user = secret.username
passwd = secret.password
entity = EntityAPI(username=user,password=passwd, server="unilever.preservica.com")
retention = RetentionAPI(username=user,password=passwd, server="unilever.preservica.com")


ref = "c66d5b8a-30dc-4319-bf44-92fcf0d834ca"
ref = "2c11f263-cdf0-4f55-a83a-9f30ceac3681"
e = entity.asset(ref)
r = retention.assignments(e)
for a in r:
    print(a)