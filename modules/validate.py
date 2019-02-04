import uszipcode

def menu(prompt, acceptable):
    response = raw_input("\n" + prompt + "\n> ")     
    while response not in acceptable or len(response) > 1 or response is "":
        print "\nThat's not an acceptable response. Please try again. "
        response = raw_input(prompt + "\n> ")
    return response


def integer(prompt):
    response = raw_input("\n" + prompt + "\n> ")
    while not response.isdigit():
        print "\nYou must enter a number. Please try again. "
        response = raw_input(prompt + "\n> ")
    response = int(response)
    return response


def boolean(prompt):
    res = raw_input("\n" + prompt + "\n> ").lower()
    acceptable = ['yes', '1', 'true', 'y', 'no', '0', 'false', 'n']
    while res not in acceptable or res is "":
        print "\nThat's not an acceptable response. Please try again. "
        res = raw_input(prompt + "\n> ")
    if res == "no" or res == "false" or res == "0" or res == "n":
        return False
    elif res == "yes" or res == "true" or res == "1" or res == "y":
        return True


def string(prompt):
    response = raw_input("\n" + prompt + "\n> ")
    return response


def zip_find(recipientType):
    code = string("Enter the " + recipientType + "'s mailing zip code.")
    code = ''.join(char for char in code if char.isdigit())
    search = uszipcode.ZipcodeSearchEngine()
    myzip = search.by_zipcode(code)
    if not myzip:
        return string("Manually enter the city, state, and zip code.")
    else:
        return myzip.City + ", " + myzip.State + " " + code


def file_len(fname):
   with open(fname) as f:
        for i, l in enumerate(f):
             pass
   return i + 1
