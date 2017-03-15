import sys


class Candidate:
    def __init__(self, name, email, subject, template, tech, status):
        self.name = name
        self.email = email
        self.subject = subject
        self.template = template
        self.tech = tech
        self.status = status


def send_email(to, subject: str, template: str):
    """Send the email to the candidates

    It sends an email using mariano.selvaggi@whiteprompt.com with the right template
    """
    import smtplib, string

    body = "\n".join(["From: %s" % gmail_user, "To: %s" % to, "Subject: %s" % subject, template])

    try:
        server = smtplib.SMTP('smtp.gmail.com')  # ,587)
        # server.ehlo()
        server.starttls()
        server.login(gmail_user, gmail_pass)
        server.sendmail(gmail_user, to, body)
        # server.close()
        server.quit()
    except Exception as e:
        raise e


def get_template(candidate):
    """Get the right template to include in the body of the email
    Get the right file and parse the information in order to create a good speech
    """

    template = ""
    try:
        with open("templates/" + candidate.template + ".txt", "r") as infile:
            template = infile.read()
            template = template.replace("[Name]", candidate.name)
            template = template.replace("[Tech]", candidate.tech)
    except FileNotFoundError:
        raise Exception('There is no file for this template')
    except Exception as e:
        raise e
    return template


def get_candidates_txt(file):
    """Get the candidates from source

    It gets the different candidates from the txt file with all the information needed
    """

    candidates = []
    try:
        with open(file, "r") as infile:
            for line in infile:
                items = line.split('|')
                status = ""
                if len(items) > 5:
                    status = str(items[5])
                candidates.append(
                    Candidate(items[0].strip(), items[1].strip(), items[2].strip(), items[3].strip(), items[4].strip(),
                              status))
    except FileNotFoundError:
        raise Exception('There is no file for this txt')
    except Exception as e:
        raise e

    return candidates


def get_candidates_excel(file):
    """Get the candidates from source

    It gets the different candidates from the excel file with all the information needed
    """
    from xlrd import open_workbook

    candidates = []

    try:
        wb = open_workbook(file)

        sheet = wb.sheets()[0]
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols

        for row in range(1, number_of_rows):
            values = []
            for col in range(0, number_of_columns):
                value = (sheet.cell(row, col).value)
                try:
                    value = str(value)
                except:
                    value = ""
                finally:
                    values.append(value)

            candidate = Candidate(name=values[0], email=values[1], subject=values[2], template=values[3],
                                  tech=values[4], status=values[5])
            candidates.append(candidate)
    except FileNotFoundError:
        raise Exception("There is no excel file")
    except Exception as e:
        raise e
    finally:
        return candidates


def mark_file_as_sent(file, candidate, row):
    """Mark the row as sent

    It marks the specific email in the txt file
    """
    if file.endswith(".txt"):
        try:
            full_line = ""
            # read the file and create changes
            with open(file, "r") as infile:
                j = 1
                for line in infile:
                    newline = line.rstrip()
                    if j == row:
                        newline += " | sent"
                    j += 1
                    total_line = full_line + newline + "\n"
            # write the file with changes
            with open(file, "w") as infile:
                infile.write(full_line.rstrip())
        except FileNotFoundError:
            raise Exception('There is no file for this txt')
        except Exception as e:
            raise e
    else:
        try:
            from xlrd import open_workbook
            from xlutils.copy import copy

            rb = open_workbook(file)
            wb = copy(rb)

            sheet = wb.get_sheet(0)
            sheet.write(row, 5, "sent")

            wb.save(file)
        except FileNotFoundError:
            raise Exception('There is no file for this txt')
        except Exception as e:
            raise e


def main(file):
    import smtplib, string

    candidates = []
    i = 1

    # Get the list of candidates from files
    try:
        if file.endswith(".txt"):
            candidates = get_candidates_txt(file)
        elif file.endswith(".xlsx") or file.endswith(".xls"):
            candidates = get_candidates_excel(file)
        else:
            print("you must input either a txt or an excel file")
    except:
        print(sys.exc_info()[0])

    # looping the candidates to send each email
    try:
        server = smtplib.SMTP('smtp.gmail.com')  # ,587)
        # server.ehlo()
        server.starttls()
        server.login(gmail_user, gmail_pass)
        for candidate in candidates:
            try:
                template = get_template(candidate)  # get the text to sent the email
                if "sent" not in candidate.status:
                    body = "\n".join(
                        ["From: %s" % gmail_from, "To: %s" % candidate.email, "Subject: %s" % candidate.subject,
                         candidate.template])
                    server.sendmail(gmail_user, candidate.email, body)
                    mark_file_as_sent(file, candidate,
                                   i)  # mark the file so the next time the same email is not sent again
                    print("new email to:" + candidate.email + " using %s" % candidate.template + "\n")
                i = i + 1
            except Exception as e:
                print("Unexpected error:", e)
                break
        server.quit()
    # server.close()
    except Exception as ex:
        print("Error in smtp:", ex)


def read_config():
    """Read the config file
    Obtain the most important key to make execute the program
    """
    import json

    data = []
    with open('config.json') as json_data_file:
        data = json.load(json_data_file)

    return data


data = read_config() #read the data from the cnofig file
gmail_user = data["mail"]["user"]
gmail_pass = data["mail"]["pass"]
gmail_from = data["mail"]["from"]

# start the process
main(input("Please enter the file name (include the file ext): "))