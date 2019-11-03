import xlrd
import os

def read_original_excel(file_name):
    data=xlrd.open_workbook(os.path.join(os.path.dirname(__file__),'res//'+file_name))
    table=data.sheet_by_index(0)
    parse_result={}
    #获得境内发货人
    sender=get_simple_value(table,'境内发货人','境外收货人','出境关别')
    parse_result['sender']=sender
    #获得境外收货人
    receiver=get_simple_value(table,'境外收货人','生产销售单位','运输方式')
    parse_result['receiver']=receiver
    return parse_result

allow_field_list=['境内发货人','境外收货人','出境关别']

#根据传入的字段标题,自动获得对应字段的值
def get_simple_field(table,field_name):
    field_value=''
    field_name_row_num=-1
    field_name_col_num=-1
    if field_name not in allow_field_list:
        field_value='不支持的字段名'
    else:
        for row_num in range(table.nrows):
            row_data_list=table.row_values(row_num)
            for col_num,data in enumerate(row_data_list):
                if str(data).find(field_name)!=-1:
                    field_name_row_num=row_num
                    field_name_col_num=col_num
                    break
        if field_name_row_num!=-1 and field_name_col_num!=-1:
            area_info=get_field_crop_area(table,field_name_row_num,field_name_col_num)
            #有下侧节点和右侧节点
            if area_info.get('have_below_field') and area_info.get('have_right_field'):
                get_field_value_according_below_right(table,area_info)
        else:
            field_value='该单据内未找到 '+field_name+' 项'
    return field_value

#通过判断获得某个字段的识别范围
def get_field_crop_area(table,field_name_row_num,field_name_col_num):
    #查找下方节点
    have_below_field=False
    field_below_row_num=-1
    for i in range(1,4):
        cell_value=table.row(field_name_row_num+i)[field_name_col_num].value
        if cell_value in allow_field_list:
            have_below_field=True
            field_below_row_num=field_name_row_num+i
            break
    #查找右侧节点
    have_right_field=False
    field_right_col_num=-1
    for i in range(1,6):
        if (field_name_col_num+i) >= len(table.row_values(field_name_row_num)):
            break
        cell_value=table.row(field_name_row_num)[field_name_col_num+i].value
        if cell_value in allow_field_list:
            have_right_field=True
            field_right_col_num=field_name_col_num+i
            break
    #查找左侧节点
    have_left_field=False
    field_left_col_num=-1
    for i in range(1,6):
        if(field_name_col_num-i<0):
            break
        cell_value=table.row(field_name_row_num)[field_name_col_num-i].value
        if cell_value in allow_field_list:
            have_left_field=True
            field_left_col_num=field_name_col_num-i
            break
    result={}
    result['have_below_field']=have_below_field
    result['field_below_row_num']=field_below_row_num
    result['have_right_field']=have_right_field
    result['field_right_col_num']=field_right_col_num
    result['have_left_field']=have_left_field
    result['field_left_col_num']=field_left_col_num
    result['field_name_row_num']=field_name_row_num
    result['field_name_col_num']=field_name_col_num
    return result

#根据下侧和右侧节点辅助进行当前内容识别
def get_field_value_according_below_right(table,area_info):
    field_name_row_num=area_info.get('field_name_row_num')
    field_name_col_num=area_info.get('field_name_col_num')
    field_below_row_num=area_info.get('field_below_row_num')
    field_right_col_num=area_info.get('field_right_col_num')
    field_value=''
    if field_below_row_num-field_name_row_num==1:
        for col_num in range(field_name_col_num,field_right_col_num):
            field_value+=table.row(field_name_col_num)[col_num].value
    else:
        pass
    return field_value

#获得境内发货人
def get_sender_value(table):
    sender_row_index=-1
    sender_col_index=-1
    receiver_row_index=-1
    receiver_col_index=-1
    leave_border_col_index=-1
    result=''
    #判断excel格式
    if(table!=None):
        for row_num in range(table.nrows):
            row_data_list=table.row_values(row_num)
            for index,data in enumerate(row_data_list):
                if(str(data).find('境内发货人')!=-1):
                    sender_row_index=row_num
                    sender_col_index=index
                if(str(data).find('境外收货人')!=-1):
                    receiver_row_index=row_num
                    receiver_col_index=index
                if(str(data).find('出境关别')!=-1):
                    leave_border_col_index=index
    if(sender_row_index!=-1 and sender_col_index!=-1 and receiver_row_index!=-1 and receiver_col_index!=-1 and leave_border_col_index!=-1):
        diff_row_num=receiver_row_index-sender_row_index
        if diff_row_num==2:
            for col_index in range(sender_col_index,leave_border_col_index):
                result+=table.row(sender_row_index+1)[col_index].value
        elif diff_row_num==1:
            for col_index in range(sender_col_index,leave_border_col_index):
                result+=table.row(sender_row_index)[col_index].value
            result=result.splitlines()[1]
    return result

def get_simple_value(table,target,below,right):
    sender_row_index=-1
    sender_col_index=-1
    receiver_row_index=-1
    receiver_col_index=-1
    leave_border_col_index=-1
    result=''
    #判断excel格式
    if(table!=None):
        for row_num in range(table.nrows):
            row_data_list=table.row_values(row_num)
            for index,data in enumerate(row_data_list):
                if(str(data).find(target)!=-1):
                    sender_row_index=row_num
                    sender_col_index=index
                if(str(data).find(below)!=-1):
                    receiver_row_index=row_num
                    receiver_col_index=index
                if(str(data).find(right)!=-1):
                    leave_border_col_index=index
    if(sender_row_index!=-1 and sender_col_index!=-1 and receiver_row_index!=-1 and receiver_col_index!=-1 and leave_border_col_index!=-1):
        diff_row_num=receiver_row_index-sender_row_index
        if diff_row_num==2:
            for col_index in range(sender_col_index,leave_border_col_index):
                result+=table.row(sender_row_index+1)[col_index].value
        elif diff_row_num==1:
            for col_index in range(sender_col_index,leave_border_col_index):
                result+=table.row(sender_row_index)[col_index].value
            result=result.splitlines()[1]
    return result



if(__name__=='__main__'):
    read_original_excel('报关单2.xlsx')
