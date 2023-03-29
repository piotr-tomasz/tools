def db_quickview(database):
    for k1 in database.keys():
        for k2 in database[k1].keys():
            try:
                print(str(k1),'-',str(k2),'-',database[k1][k2]['description'])
            except:
                pass