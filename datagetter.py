try:
    from urllib.request import Request, urlopen
except ImportError:
    from urllib2 import Request, urlopen
import os
import sys
import datetime
import time
import glob
import json
import re
import xlsxwriter

MAGIC_URL_FILE = 'url.txt'
XLS_FILE = 'results.xlsx'
DATA_FOLDER = 'data'
RETRY_COUNT = 15
SLEEP_TIME = 24 * 60 * 60
CHARTS_PER_SHEET = 5

DAILY_HOUR = 1800

APT_TYPES = {'S1': {'idx': 0, 'name': 'Studio'},
             '11': {'idx': 1, 'name': '1Bd1Ba'},
             '22': {'idx': 2, 'name': '2Bd2Ba'}}


def readTheMagicUrl():
    with open(MAGIC_URL_FILE, 'rt') as urlfile:
        return urlfile.read().strip()


def readTheData(url):
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64)'}
    exception = None
    for idx in range(0, RETRY_COUNT + 1):
        try:
            req = Request(url, None, headers)
            response = urlopen(req)
            return (response.read(), None)
        except Exception as e:
            exception = e
            print('Error, retrying (%d/%d)...' % (idx + 1, RETRY_COUNT))
            time.sleep(2)
    return None, exception


def toPyDate(dateStr):
    if '-' in dateStr:
        return datetime.datetime.strptime(dateStr, '%Y-%m-%d')
    return datetime.datetime.strptime(dateStr, "%m/%d/%Y")


def retrieveData():
    magic_url = readTheMagicUrl()
    crt_date = datetime.datetime.now()
    file_name = crt_date.strftime('%Y-%m-%d.json')
    full_path = os.path.join(DATA_FOLDER, file_name)
    if not os.path.exists(full_path):
        print('Data not found locally, retrieving...')
        data, exception = readTheData(magic_url)
        if data is None:
            print ('Could not retrieve data: ' + str(exception))
        else:
            try:
                with open(full_path, 'wt') as output:
                    output.write(data)
                print('Wrote ' + str(len(data) >> 10) + ' KB')
                return True
            except Exception as e:
                print('Could not write data: ' + str(e))
    else:
        print('Data found locally, skipping...')
    return False


def processFile(file_path, processed_data):
    json_data = None
    with open(file_path, 'rt') as input:
        string_data = input.read()
    string_data = re.sub(",[ \t\r\n]*}", "}", string_data)
    string_data = re.sub(",[ \t\r\n]*]", "]", string_data)
    json_data = json.loads(string_data)
    results = json_data['resultsets'][0]['results']
    date_string = os.path.basename(file_path).split('.')[0]
    nb_results = len(results)
    nb_filtered = 0
    processed_date = toPyDate(date_string)

    results_data = processed_data['results_data']
    results_count = processed_data['results_count']
    processed_data['processed_dates'].append(processed_date)

    for apt_type in APT_TYPES:
        results_count[apt_type].append(0)

    for result in results:
        apt_code = result['unitTypeCode'][:2]
        if apt_code not in APT_TYPES:
            nb_filtered += 1
            continue
        results_count[apt_code][-1] += 1
        building_name = result['buildingName']
        unit_name = result['unitMarketingName']
        full_name = building_name + '-' + unit_name
        if full_name not in results_data[apt_code]:
            ap_info = {}
            ap_info['marketRent'] = [float(result['marketRent'])]
            ap_info['buildingName'] = result['buildingName']
            ap_info['unitSqFt'] = int(result['unitSqFt'])
            ap_info['floorplanMarketingName'] = result[
                                                  'floorplanMarketingName']
            ap_info['unitMarketingName'] = result['unitMarketingName']
            ap_info['unitPricingDate'] = [toPyDate(result['unitPricingDate'])]
            ap_info['unitBestPrice'] = [float(result['unitBestPrice'])]
            ap_info['unitBestDate'] = [toPyDate(result['unitBestDate'])]
            ap_info['unitBestTerm'] = [int(result['unitBestTerm'])]
            ap_info['sampleDates'] = [processed_date]
            results_data[apt_code][full_name] = ap_info
        else:
            ap_info = results_data[apt_code][full_name]
            if result['unitPricingDate'] != ap_info['unitPricingDate'][-1]:
                ap_info['unitPricingDate'].append(toPyDate(
                                                    result['unitPricingDate']))
            ap_info['marketRent'].append(float(result['marketRent']))
            ap_info['unitBestPrice'].append(float(result['unitBestPrice']))
            ap_info['unitBestDate'].append(toPyDate(result['unitBestDate']))
            ap_info['unitBestTerm'].append(int(result['unitBestTerm']))
            ap_info['sampleDates'].append(processed_date)


def postProcessData(processed_data):
    for apt_type in APT_TYPES:
        apt_data = processed_data['results_data'][apt_type]
        for ap_name, apartment in apt_data.items():
            apartment['avgBestPrice'] = (sum(apartment['unitBestPrice']) /
                                         len(apartment['unitBestPrice']))
            apartment['avgMarketRent'] = (sum(apartment['marketRent']) /
                                          len(apartment['marketRent']))
            apartment['state'] = 'normal'
            if (apartment['sampleDates'][-1] <
                    processed_data['processed_dates'][-1]):
                apartment['state'] = 'expired'
            elif len(apartment['sampleDates']) == 1:
                apartment['state'] = 'new'
                processed_data['new_items'][apt_type] = True


def processDataAndGenerateXLS():
    print('Processing data files')
    processed_data = {'results_data': {apt: {} for apt in APT_TYPES},
                      'results_count': {apt: [] for apt in APT_TYPES},
                      'new_items': {apt: False for apt in APT_TYPES},
                      'processed_dates': []}
    files = sorted(glob.glob(os.path.join(DATA_FOLDER, '*.json')))
    for idx, file in enumerate(files):
        print('Processing file %d/%d' % (idx + 1, len(files)))
        processFile(file, processed_data)
        print('Done')
    postProcessData(processed_data)
    print('Writing XLS file')
    workbook = xlsxwriter.Workbook(XLS_FILE)
    title_format = workbook.add_format()
    title_format.set_font_size(14)
    title_format.set_bold()
    exp_format = workbook.add_format()
    exp_format.set_font_size(14)
    exp_format.set_bold()
    exp_format.set_font_color('#e6003a')
    exp_format.set_bg_color('#040404')
    new_format = workbook.add_format()
    new_format.set_font_size(14)
    new_format.set_bold()
    new_format.set_font_color('#daa520')
    new_format.set_bg_color('#040404')
    info_format = workbook.add_format()
    info_format.set_font_color('blue')
    date_format = workbook.add_format()
    date_format.set_font_color('#bf6900')
    date_format.set_num_format("MM/DD/YYYY")
    important_date_format = workbook.add_format()
    important_date_format.set_font_size(12)
    important_date_format.set_bold()
    important_date_format.set_font_color('#fc521e')
    important_date_format.set_num_format("MM/DD/YYYY")
    important_format = workbook.add_format()
    important_format.set_font_size(12)
    important_format.set_bold()
    important_format.set_font_color('#fc521e')
    sorted_apt_types = sorted(APT_TYPES.keys(),
                              key=lambda x: APT_TYPES[x]['idx'])

    def w(row, col, val, format=None):
        worksheet.write(row, col, val, format)

    def wr(row, col, val, format=None):
        worksheet.write_row(row, col, val, format)

    for apt_type in sorted_apt_types:
        apt_data = processed_data['results_data'][apt_type]
        sheet_name = APT_TYPES[apt_type]['name']
        if processed_data['new_items'][apt_type]:
            sheet_name += '(!)'
        worksheet = workbook.add_worksheet(sheet_name)
        apartment_names = sorted(apt_data.keys(),
                                 key=lambda x:
                                 apt_data[x]['unitBestPrice'][-1])
        crt_row = 0
        worksheet.set_column(0, 80, 25)
        for apt_idx, name in enumerate(apartment_names):
            apt = apt_data[name]
            start_row = crt_row
            w(crt_row, 0, 'Apartment:', title_format)
            w(crt_row, 1, name)
            if apt['state'] == 'expired':
                w(crt_row, 2, 'EXPIRED', exp_format)
            elif apt['state'] == 'new':
                w(crt_row, 2, 'NEW', new_format)
            crt_row += 1
            w(crt_row, 0, 'Square footage:')
            w(crt_row, 1, apt['unitSqFt'], info_format)
            crt_row += 1
            w(crt_row, 0, 'Current Best Price:', important_format)
            w(crt_row, 1, apt['unitBestPrice'][-1], important_format)
            w(crt_row, 2, round(apt['unitBestPrice'][-1] /
              apt['unitSqFt'], 3), info_format)
            w(crt_row, 3, 'per sq foot', info_format)
            crt_row += 1
            w(crt_row, 0, 'Avg Best Price:')
            w(crt_row, 1, round(apt['avgBestPrice'], 3), info_format)
            w(crt_row, 2, round(apt['avgBestPrice'] /
              apt['unitSqFt'], 3), info_format)
            w(crt_row, 3, 'per sq foot', info_format)
            crt_row += 1
            w(crt_row, 0, 'Avg Market Price:')
            w(crt_row, 1, round(apt['avgMarketRent'], 3),
              info_format)
            w(crt_row, 2, round(apt['avgMarketRent'] /
              apt['unitSqFt'], 3), info_format)
            w(crt_row, 3, 'per sq foot', info_format)
            crt_row += 1
            w(crt_row, 0, 'Start date:', important_format)
            w(crt_row, 1, apt['unitBestDate'][-1],
              important_date_format)
            crt_row += 1
            w(crt_row, 0, 'Pricing date:')
            w(crt_row, 1, apt['unitPricingDate'][-1], date_format)
            crt_row += 4
            w(crt_row, 0, 'Dates:')
            wr(crt_row, 1, apt['sampleDates'], date_format)
            dates_row = crt_row
            crt_row += 1
            w(crt_row, 0, 'Best prices:')
            wr(crt_row, 1, apt['unitBestPrice'], info_format)
            prices_row = crt_row
            crt_row += 1
            w(crt_row, 0, 'Lease length:')
            wr(crt_row, 1, apt['unitBestTerm'], info_format)
            crt_row += 1
            w(crt_row, 0, 'Market value:')
            wr(crt_row, 1, apt['marketRent'], info_format)
            market_row = crt_row
            crt_row += 1
            if apt_idx < CHARTS_PER_SHEET:
                chart = workbook.add_chart({'type': 'line'})
                chart.add_series({
                    'categories': [sheet_name,
                                   dates_row, 1,
                                   dates_row,
                                   len(apt['sampleDates'])],
                    'values':     [sheet_name,
                                   prices_row, 1,
                                   prices_row,
                                   len(apt['sampleDates'])],
                    'line':       {'color': 'orange'},
                    'name':       'RV',
                    'data_labels': {'value': True},
                })
                chart.add_series({
                    'categories': [sheet_name,
                                   dates_row, 1,
                                   dates_row,
                                   len(apt['sampleDates'])],
                    'values':     [sheet_name,
                                   market_row, 1,
                                   market_row,
                                   len(apt['sampleDates'])],
                    'line':       {'color': 'blue'},
                    'name':       'MK',
                    'data_labels': {'value': True},
                })
                chart.set_legend({'none': True})
                worksheet.insert_chart(start_row, 4, chart)
            crt_row += 5
    sheet_name = 'Summary'
    worksheet = workbook.add_worksheet(sheet_name)
    worksheet.set_column(0, 80, 25)
    crt_row = 0
    w(crt_row, 0, 'Availability', title_format)
    crt_row += 2
    dates_row = crt_row
    w(crt_row, 0, "Dates:")
    wr(crt_row, 1, processed_data['processed_dates'], date_format)
    crt_row += 1
    for apt_type in sorted_apt_types:
        w(crt_row, 0, APT_TYPES[apt_type]['name'])
        wr(crt_row, 1, processed_data['results_count'][apt_type], info_format)
        crt_row += 1
    chart = workbook.add_chart({'type': 'line'})
    for apt_idx, apt_type in enumerate(sorted_apt_types):
        values_row = dates_row + apt_idx + 1
        chart.add_series({
            'categories': [sheet_name,
                           dates_row, 1,
                           dates_row,
                           len(processed_data['processed_dates'])],
            'values':     [sheet_name,
                           values_row, 1,
                           values_row,
                           len(processed_data['processed_dates'])],
            'name':       '=' + sheet_name + '!$A$' + str(values_row + 1),
            'data_labels': {'value': True},
        })
    worksheet.insert_chart(crt_row, 0, chart)
    workbook.close()
    print('Done writing XLS file')


def getCrtTimeMilitary():
    t = datetime.datetime.now()
    return t.hour * 100 + t.minute


def main(forced):
    print('Data extractor v001\n')
    last_checked = getCrtTimeMilitary()
    while True:
        current_time = getCrtTimeMilitary()
        if (last_checked < DAILY_HOUR and
                current_time >= DAILY_HOUR) or forced:
            print('')
            retrieveData()
            processDataAndGenerateXLS()
            forced = False
            print('')
        last_checked = current_time
        remaining_h = 0
        remaining_m = 0
        hours_c = current_time // 100
        hours_t = DAILY_HOUR // 100
        if current_time >= DAILY_HOUR:
            hours_t += 24
        mins_c = current_time % 100
        mins_t = DAILY_HOUR % 100
        remaining_h = hours_t - hours_c
        remaining_m = mins_t - mins_c
        if remaining_m < 0:
            remaining_m += 60
            remaining_h -= 1
        sys.stdout.write('\r%02d:%02d until the next update' %
                         (remaining_h, remaining_m))
        sys.stdout.flush()
        time.sleep(60)


if __name__ == '__main__':
    main(len(sys.argv) > 1)
