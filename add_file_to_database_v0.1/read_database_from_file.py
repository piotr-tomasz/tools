import json
def read_database_from_file(path):
    f=open(path)
    measurements=json.load(f)
    f.close()
    return(measurements)