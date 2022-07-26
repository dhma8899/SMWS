import json

def trial(name):
    f=open('Database.json')
    data=json.load(f)
    return data[name]







