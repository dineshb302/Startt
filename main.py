from fastapi import Request, FastAPI
import func as f
import boto3
import s3fs
import datetime
import uvicorn
import re
import uuid
import os
import time
from config import *

s3 = boto3.client('s3')

aws_access_key_id = access_id
aws_secret_access_key = access_key

f.os.environ["AWS_DEFAULT_REGION"] = 'ap-south-1'
f.os.environ["AWS_ACCESS_KEY_ID"] = aws_access_key_id
f.os.environ["AWS_SECRET_ACCESS_KEY"] = aws_secret_access_key
s3 = boto3.resource(
    service_name='s3',
    region_name='ap-south-1',
    aws_access_key_id = aws_access_key_id,
    aws_secret_access_key= aws_secret_access_key
)


def action(data_row):
    timestamp = data_row[0]
    date = data_row[0][:10] 
    name = data_row[1]
    income_as_input = data_row[12]
    income = f.income_and_saving(data_row)[0]
    income_source = data_row[13]
    savings_as_input = data_row[14]
    savings = f.income_and_saving(data_row)[1]
    life_stage = data_row[16] 
    age = data_row[18]
    net_worth_as_input = data_row[15]
    income_exact =  round(float(re.sub('\s*lakhs*', '', data_row[19]))*10**5)                        # int(data_row[19]*10**5)
    mobile = data_row[20] 
    sip = savings*income_exact
    namef = name.replace(' ', '')
    code = timestamp.replace('/','').replace(':','').replace(' ','')
    
    data_dict = f.nest()
    
    aror = 0.10
    niftyif, itif, scif, ldif = 0.15, 0.2846, 0.1199, 0.07
    
    sum_capacities, sum_tolerances, avg_capacities, avg_tolerances, total, avg_total = f.scores(data_row, data_dict)
    debt, equity, longDurationBond, nifty, it, smallCap= f.risk_profile(avg_total)
    nif = (niftyif * nifty)/equity + (itif*it)/equity + (smallCap*scif)/equity

    figi = (ldif*longDurationBond)/debt

    invested_amount, portfolio_amount, years = f.schedule(sip, debt, equity, nif, figi)
    df_forecast = f.forecast(invested_amount, portfolio_amount, years, sip, debt, equity, nif, figi) 
    remarks = f.tolerance_remarks(avg_capacities, avg_tolerances)



    f.stackbar(round(debt,1), round(equity,1), name, mobile, code)    
    f.line_chart(df_forecast, name, mobile, code)
    f.gauge(name, mobile, code,labels=['Conservative','Moderate','Balanced','Assertive','Aggressive'], colors=["#FFB6C1","#EE6363","#CD5555","#8B3A3A","#800000"], arrow=f.pointer(avg_total), size=(5,3), title=str('Your Risk Score is {}'.format(int(avg_total))))

    f.pptx_work("Investor Wealth Report-v3.pptx", name, mobile, code, date, 
                sip, aror, figi, avg_capacities, avg_tolerances, 
                net_worth_as_input, age, life_stage, savings_as_input, 
                income_source, income_as_input, data_row,longDurationBond,
                nifty, it, smallCap, nif)
    
    f.to_pdf('./','base.pptx',name,code)
    

    
    
    #s3.Bucket('starttbucket').upload_file(Filename=str(name)+str(code)+'.pdf', Key=str(name)+str(code)+'.pdf')
    s3.Bucket('starttbucket').upload_file(Filename='base.pdf', Key=str(name)+str(code)+'.pdf')
    time.sleep(2)
    url = boto3.client('s3').generate_presigned_url('get_object', Params = {'Bucket': 'starttbucket','Key': str(name) + str(code) + '.pdf'}, ExpiresIn = 100000)
    return url
   #get public url of this file and return

app = FastAPI()


@app.post("/analysis")
async def main(request: Request):

    data = await request.json()
    datetime_object = datetime.datetime.now()
    # print(data)
    data_list = [
        str(datetime_object).split(".")[0],
        data['name']
    ]
    for index in range(1,17):
        key = "q" + str(index)
        data_list.append(data[key])

    data_list.append(data['age'])
    data_list.append(data['income'])
    data_list.append(data['phone'])
    url = action(data_list)     # assign url to variable like --> url = action(data_list)
    file_name = str(uuid.uuid4())
    cmd = "mv base.pdf {}.pdf".format(file_name)
    os.system(cmd)
    return {"res": True, "url": url}  # return url with response --> {"res": True, "url": url}

if __name__ == '__main__':
    # workers --> to handle multiple user requests at a time
    uvicorn.run('main:app', host='0.0.0.0', port=5007, workers=1)