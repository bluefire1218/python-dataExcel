# -*- coding: gbk -*-
import requests
import xlsxwriter

class parsingApi:
    # ���е���Ϣ
    @property           ##�����������ԣ����Ƿ����� ����ʱ��ֱ�ӵ��÷�������
    def information(self):
        url = 'http://pi.vaiwan.com/piwebapi/'
        res = requests.get(url)
        # print("---res��{}".format(res.json()))
        data = res.json()['Links']['DataServers']
        da = data.split('/')[-1]

        url = 'http://pi.vaiwan.com/piwebapi/'
        path = '{}{}'.format(url, da)
        res_list = requests.get(path)
        # print("---res_list��{}".format(res_list.json()))
        data = res_list.json()['Items'][0]['Links']['Points']
        # print("---data: {}".format(data))
        da = data.split('dataservers/')[1]
        # print("---da: {}".format(da))

        url = 'http://pi.vaiwan.com/piwebapi/dataservers/'
        path = '{}{}'.format(url, da)
        some_list = requests.get(path)      ##��ȡ�����еĵ���Ϣ
        # print("---some_list: {}".format(some_list.json()))

        results_arr = []                    # ����һ��list�����������

        for item in some_list.json()['Items']:   ## ѭ������Ϣ
            name = item['Name']                     ## �������name
            point_type = item['PointType']          ## ������� type
            record_data = item['Links']['RecordedData']     #���InterpolatedData
            links = record_data.split('/streams/')[1]
            url = 'http://pi.vaiwan.com/piwebapi/streams/{}'.format(links)  ## ƴ�Ӻ���ÿ��name��Ӧ��url

            stream_datas = requests.get(url).json()     ## ѭ������ÿ��url

            values = []
            for v in stream_datas['Items']:     ##ѭ��ÿ��name�����url�������
                timestamp = v['Timestamp']      ## ʱ��ֱ�ӻ�ȡ
                good = v['Good']                ## goodֱ�ӻ�ȡ
                value = 0
                if isinstance(v['Value'], dict): ##�ж������url�е� value �ǲ����ֵ�����
                    if v['Value']['Value']:         ##������ֵ�����ȡvalue���µ�value����ֵ
                        value = v['Value']['Value']     ## valueȡֵ
                else:
                    value = v['Value']      ##value�����ֵ䣬ֱ��ȡֵ

                r = [timestamp, value, good]           ## ��name ��Ӧ��һ�� ʱ�䣬value��good ����һ��list
                # print("r��ֵ----->",r)
                values.append(r)                       ## ��ÿһ��name ȡ�õ� ʱ�䣬value��good ����һ��list
                # print("values��ֵ----->",values)

            row_dict = {'name': name, 'point_type': point_type, 'values': values}    ##�� һ�������Ϣ name ,���ͣ� ���ʱ�䣬value��good��list  ȫ�� �����ֵ�
            # print("row_dict��ֵ----->",row_dict)
            results_arr.append(row_dict)                            ## �����ÿһ�������Ϣ�� �ֵ� ����list
        # print("result_arr��ֵ----->",results_arr)

        return results_arr

    # д��
    def write_excel(self,datas,file_path):
        print("------excelд�����ݣ�{}".format(datas))
        print("------excelд���ļ���{}".format(file_path))

        workbook = xlsxwriter.Workbook('{}'.format(file_path))  # �����ļ�
        worksheet = workbook.add_worksheet()  # ����sheet

        worksheet.write(0, 0, '{}'.format("name"))
        worksheet.write(0, 1, '{}'.format("point_type"))
        worksheet.write(0, 2, '{}'.format("time"))
        worksheet.write(0, 3, '{}'.format("count"))
        worksheet.write(0, 4, '{}'.format("yesOrNo"))

        temp = 0

        for index, item in enumerate(datas):
            index = index + 1

            if len(item['values']) == 0:
                index = index + temp
                worksheet.write(index, 0, '{}'.format(item['name']))
                worksheet.write(index, 1, '{}'.format(item['point_type']))
                worksheet.write(index, 2, '{}'.format(item['values']))
            else:
                for ind, it in enumerate(item['values']):
                    tm = index
                    if ind != 0:
                        temp = temp + 1
                    index = index + temp
                    print("---it�� {}".format(it))
                    print("---temp�� {}".format(temp))
                    print("---index�� {}".format(index))

                    worksheet.write(index, 0, '{}'.format(item['name']))
                    worksheet.write(index, 1, '{}'.format(item['point_type']))
                    worksheet.write(index, 2, '{}'.format(str(it[0])))
                    worksheet.write(index, 3, '{}'.format(str(it[1])))
                    worksheet.write(index, 4, '{}'.format(str(it[2])))

                    index = tm

        workbook.close()



if __name__ == '__main__':
    parsingApi().write_excel(parsingApi().information,'excel/data.xlsx')