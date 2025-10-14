def harm_codes(fp):        # returns the contents of a text file as a list of strings from a file path provided as a string
    codes = []
    with open(fp) as f:
        file_contents = f.read()    
    for content in file_contents.split():
        codes.append(content)
    return codes
