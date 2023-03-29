def sort_dictionary(input_dictionary):
    '''Function sorts first level of dictionary keys'''
    input_dictionary=dict(sorted(input_dictionary.items(), key=lambda item: item[0]))
    return input_dictionary