# this package is for handling email  filter

class EmailsHandler:
    __conditions = []
    __sptripEmails = []

    def getDataRequiredSheet(self, emaillist = []):
        self.__conditions = emaillist
        self.__emailsRecieved()


    def __emailsRecieved(self):

        try:
            for email in self.__conditions:
                stmp = email.split('@')[-1].lower()
                self.__sptripEmails.append(stmp)
        except:
            pass


    def getEmailStmp(self):
        return self.__sptripEmails


    def toLowerEmails(self,emails):

        all = []
        for email in emails:
            all.append(email.lower())
        return all

