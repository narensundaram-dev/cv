class EmailNormalizer(object):

    def __init__(self, email):
        self.email = email

    def trim(self, return_=False):
        new_string = ""
        found = False
        for s in self.email:
            if not found:
                if s.isalpha():
                    new_string += s
                    found = True
            else:
                new_string += s
        if return_:
            return new_string[::-1]
        self.email = new_string[::-1]
        return self.trim(return_=True)

    def normalize(self):
        return self.trim()


if __name__ == '__main__':
    print(EmailNormalizer("../naren@gmail.com").normalize())
    print(EmailNormalizer("naren@gmail.com12..,3").normalize())
    print(EmailNormalizer("../naren@gmail.com12..,3").normalize())
