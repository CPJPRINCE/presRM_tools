from pyPreservica import *
from presRM_labels.labels.options import BOX_SEARCH_FILTERS,FILE_SEARCH_FILTERS
from secret import *

CONTENT = ContentAPI(username,password,server)
ENTITY = EntityAPI(username,password,server)

def login_preservica(username=None,password=None,server=None):
    if username == None or password == None:
        print('Username or Password is set to none... Please address...')
        raise SystemExit()
    print('Successfully logged into Preservica')

def search_preservica(**kwargs):
    search = list(CONTENT.search_index_filter_list(query="%",filter_values=kwargs))
    return search