import readEmail
from readEmail import CheckMailer


def main(daysOfReport):
    chk=CheckMailer(daysOfReport)
    chk.constructReportData()



if __name__=="__main__":
    import argparse

    parser = argparse.ArgumentParser()
    # parser.add_argument("daysOfReport=0", help="Mails of last Number of days ", type=int)
    parser.add_argument("daysOfReport", nargs='?',type=int)
    args = parser.parse_args()
    if  args.daysOfReport == None:
        main(0)
    else:
        main(args.daysOfReport)