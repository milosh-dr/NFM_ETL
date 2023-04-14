from extract import get_links, lineup_parser, docx_parser2
from test import xlsx_parser

# Xlsxy nie działają poprawnie

# Set the home url
url = r'http://www.nfm.home.pl/orkiestra/T%20A%20B%20L%20I%20C%20A%20%20%20O%20G%20%C5%81%20O%20S%20Z%20E%20%C5%83/'
# Set the credentials for authetication
username = 'muzyk'
password = 'Coda2019!'

def main():
    weeks = get_links(url, username, password)
    for i, week in enumerate(weeks):
        # lineup = lineup_parser(week['lineup'][0], username, password)
        # if i not in [0,2,8]:
        #     continue
        print(i)
        for schedule in week['schedule']:
            # print(schedule)
            if schedule.endswith('xlsx'):
                text, df = xlsx_parser(schedule, username, password)
                print(df)
            elif schedule.endswith('docx'):
                pass
                # df = docx_simple(schedule, username, password)
                # name, artists, programme, df = docx_parser2(schedule, username, password)
            
        print('-'*30)
        # if i==5:
        #     break
    print('Done')


if __name__ == '__main__':
    main()