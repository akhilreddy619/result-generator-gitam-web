from flask import Flask, render_template,request,send_file
import pandas
import requests
from bs4 import BeautifulSoup
import os
import re
import xlsxwriter
app=Flask(__name__)

@app.route('/')
def home():
    return render_template("home.html")

@app.route('/results/',methods = ['POST'])
def akhil():
    global file_name,file_name1,final_result
    if request.method == 'POST':
        join = request.form['rollno']
        sem = request.form['semester']
        sec = request.form['section']
        payload = {
        "__EVENTTARGET":"",
        "__EVENTARGUMENT":"",
        "__VIEWSTATE":"/wEPDwULLTE3MTAzMDk3NzUPZBYCAgMPZBYCAgcPDxYCHgRUZXh0ZWRkZKKjA/8YeuWfLRpWAZ2J1Qp0eXCJ",
        "__VIEWSTATEGENERATOR":"65B05190",
        "__EVENTVALIDATION":"/wEWFAKj/sbfBgLnsLO+DQLIk+gdAsmT6B0CypPoHQLLk+gdAsyT6B0CzZPoHQLOk+gdAt+T6B0C0JPoHQLIk6geAsiTpB4CyJOgHgLIk5weAsiTmB4CyJOUHgKL+46CBgKM54rGBgK7q7GGCLOsGLAxgUwycOU5mDizjY4EVXof",
        "cbosem":"1",
        "txtreg":"1210314401",
        "Button1":"Get Result"
        }
        try:
            base = join[0:8]
            try:
                payload['cbosem'] = sem
                result = []
                for roll in range(1,68):
                    try:
                        if 1<=roll<=9:
                            payload['txtreg']=base+"0"+str(roll)
                        else:
                            payload['txtreg']=base+str(roll)
                        res = requests.post("https://doeresults.gitam.edu/onlineresults/pages/Newgrdcrdinput1.aspx", data=payload)
                        soup = BeautifulSoup(res.text,"html.parser")
                        name = soup.find("span",{"id":"lblname"}).text
                        reg = soup.find("span",{"id":"lblregdno"}).text
                        heads=[]
                        heads.append("Name")
                        heads.append("Roll No")

                        sgpa=soup.find("span",{"id":"lblgpa"}).text
                        cgpa=soup.find("span",{"id":"lblcgpa"}).text
                        table=soup.find("table",{"class":"table-responsive"})
                        rows=table.find_all("tr")[1:]
                        temp = []
                        temp.append(name)
                        temp.append(reg)
                        for row in rows:
                            count=0
                            for i in row.findAll("td"):
                                if count==3:
                                    temp.append(i.text)
                                elif count==0:
                                    z = i.text
                                elif count==1:
                                    heads.append(i.text+"("+z+")")
                                count=count+1
                        temp.append(sgpa)
                        temp.append(cgpa)
                        result.append(temp)
                    except:
                        pass
                heads.append("SGPA")
                heads.append("CGPA")
                df=pandas.DataFrame(result,index=None)
                num = join[5:7]
                n = int(num)+4
                file_name = "Year("+num+"-"+str(n)+")Sec-"+sec+"-Sem-"+sem+"-Results.csv"
                file_name1 = "Year("+num+"-"+str(n)+")Sec-"+sec+"-Sem-"+sem+"-Results.xlsx"

                df.to_csv("uploads/"+file_name,header= heads,index=False)
                final_result = "result_"+file_name1
                workbook = xlsxwriter.Workbook("uploads/"+final_result)
                worksheet = workbook.add_worksheet()
                df1=pandas.read_csv("uploads/"+file_name)
                df2=df1.iloc[:,2:-2]
                df3=df2.replace({'O':10,'A+':9,'A':8,'B+':7,'B':6,'C':5,'D':4,'F':0})
                x=df2.columns.values
                q = ['A10','J10','S10','AB10','A27','J27','S27','AB27','AK27']
                h = ['B1','E1','H1','K1','N1','Q1','T1','W1','Z1','AC1','AF1']
                k = ['B2','E2','H2','K2','N2','Q2','T2','W2','Z2','AC2','AF2']
                o = ['B','E','H','K','N','Q','T','W','Z','AC','AF']
                gt = ['A','D','F','G','J','M','P','S','V','Y','AB','AE']
                u = ['A2','D2','F2','G2','J2','M2','P2','S2','V2','Y2','AB2','AE2']
                ux = ['A1','D1','F1','G1','J1','M1','P1','S1','V1','Y1','AB1','AE1']
                cv = ["Grade"]
                for t in range(len(x)):
                    df4=df2.iloc[:,t].value_counts()

                    l=df4.to_dict()
                    y=[]
                    for i in l.keys():
                        y.append(i)

                    z=[]
                    for i in l.values():
                        z.append(i)

                # Create a new Chart object.
                    chart = workbook.add_chart({'type': 'column'})
                #Other chart commands

                #Writing data to different columns for different charts
                    worksheet.write(ux[t],cv[0])
                    worksheet.write_column(u[t], y)
                    worksheet.write(h[t],x[t])
                    worksheet.write_column(k[t], z)



                # Configure the chart. In simplest case we add one or more data series.
                #chart.add_series({'values': '=Sheet1!$A$1:$A$5'})
                    chart.add_series({'name':x[t],'categories': '=Sheet1!$'+gt[t]+'$2:$'+gt[t]+'$8','values': '=Sheet1!$'+o[t]+'$2:$'+o[t]+'$8'})


                    worksheet.insert_chart(q[t], chart)

                workbook.close()

                return render_template("result.html",btn = 'download.html')
            except:
                return "Check Again 1!"
        except:
            return "Check Again 2!"



@app.route('/download-result/')
def download():
    return send_file("uploads/"+file_name,attachment_filename=file_name,as_attachment=True)

@app.route('/download-graph/')
def download1():
    return send_file("uploads/"+final_result,attachment_filename=final_result,as_attachment=True)

if __name__=="__main__":
    app.run(debug=True)
