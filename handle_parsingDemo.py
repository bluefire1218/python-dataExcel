# -*- coding: gbk -*-
import requests
import xlsxwriter

class parsingApi:
    # 所有点信息
    @property           ##被声明是属性，不是方法， 调用时可直接调用方法本身
    def information(self):
        url = 'http://pi.vaiwan.com/piwebapi/'
        res = requests.get(url)
        # print("---res：{}".format(res.json()))
        data = res.json()['Links']['DataServers']
        da = data.split('/')[-1]

        url = 'http://pi.vaiwan.com/piwebapi/'
        path = '{}{}'.format(url, da)
        res_list = requests.get(path)
        # print("---res_list：{}".format(res_list.json()))
        data = res_list.json()['Items'][0]['Links']['Points']
        # print("---data: {}".format(data))
        da = data.split('dataservers/')[1]
        # print("---da: {}".format(da))

        url = 'http://pi.vaiwan.com/piwebapi/dataservers/'
        path = '{}{}'.format(url, da)
        some_list = requests.get(path)      ##获取到所有的点信息
        # print("---some_list: {}".format(some_list.json()))

        results_arr = []                    # 创建一个list存放所有数据

        for item in some_list.json()['Items']:   ## 循环点信息
            name = item['Name']                     ## 获得所有name
            point_type = item['PointType']          ## 获得所有 type
            record_data = item['Links']['RecordedData']     #获得InterpolatedData
            links = record_data.split('/streams/')[1]
            url = 'http://pi.vaiwan.com/piwebapi/streams/{}'.format(links)  ## 拼接后获得每个name对应的url

            stream_datas = requests.get(url).json()     ## 循环访问每个url

            values = []
            for v in stream_datas['Items']:     ##循环每个name请求的url后的数据
                timestamp = v['Timestamp']      ## 时间直接获取
                good = v['Good']                ## good直接获取
                value = 0
                if isinstance(v['Value'], dict): ##判断请求的url中的 value 是不是字典类型
                    if v['Value']['Value']:         ##如果是字典类型取value键下的value键的值
                        value = v['Value']['Value']     ## value取值
                else:
                    value = v['Value']      ##value不是字典，直接取值

                r = [timestamp, value, good]           ## 将name 对应的一组 时间，value，good 存入一个list
                # print("r的值----->",r)
                values.append(r)                       ## 将每一个name 取得的 时间，value，good 放入一个list
                # print("values的值----->",values)

            row_dict = {'name': name, 'point_type': point_type, 'values': values}    ##将 一个点的信息 name ,类型， 存放时间，value，good的list  全部 存入字典
            # print("row_dict的值----->",row_dict)
            results_arr.append(row_dict)                            ## 将存放每一个点的信息的 字典 放入list
        # print("result_arr的值----->",results_arr)

        return results_arr

    # 写入
    def write_excel(self,datas,file_path):
        print("------excel写入数据：{}".format(datas))
        print("------excel写入文件：{}".format(file_path))

        workbook = xlsxwriter.Workbook('{}'.format(file_path))  # 建立文件
        worksheet = workbook.add_worksheet()  # 建立sheet

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
                    print("---it： {}".format(it))
                    print("---temp： {}".format(temp))
                    print("---index： {}".format(index))

                    worksheet.write(index, 0, '{}'.format(item['name']))
                    worksheet.write(index, 1, '{}'.format(item['point_type']))
                    worksheet.write(index, 2, '{}'.format(str(it[0])))
                    worksheet.write(index, 3, '{}'.format(str(it[1])))
                    worksheet.write(index, 4, '{}'.format(str(it[2])))

                    index = tm

        workbook.close()



if __name__ == '__main__':
    parsingApi().write_excel(parsingApi().information,'excel/data.xlsx')