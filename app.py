from flask import Flask, render_template, request
import pandas
from fileinput import filename
import pandas as pd
import os
import openai
from tabulate import tabulate

openai.api_key = "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"

app = Flask(__name__)

@app.get('/')
def upload():
    return render_template('upload-excel.html')
    
@app.post('/view')
def view():

    file = request.files['file']
    # file_lead = pd.read_excel(file, sheet_name="Lead")
    df2 = pd.read_excel(file, sheet_name="Task")
    df3 = pd.read_excel(file, sheet_name="Event")
    
    
    #dropping date columns as they are same for all cases hence creating no variance for our model
    df3 = df3.drop(["Start Date","End Date"],1)
    #renaming few columns to merge the data
    df2 = df2.rename(columns={"TaskId": "Task_event_id"})
    df3 = df3.rename(columns={"Event ID": "Task_event_id"})
    df3 = df3.rename(columns={"Title": "Subject"})
    
    #concatincating data
    final_df = pd.concat([df2, df3])
    
    
    print(final_df.head())
    print(final_df.shape)
    
    #model name
    #text-davinci-003
    #gpt-3.5-turbo
    #\n\naverage score at lead level?
    
    #dataframe to pretty table text
    final_df_tabulate = tabulate(final_df, headers='keys', tablefmt='psql')
    #open ai
    response = openai.Completion.create(
    model="text-davinci-003",
    prompt=f"{final_df_tabulate}   \n\naverage scores of leads?",
    temperature=0.5,
    max_tokens=435,
    top_p=0.3,
    frequency_penalty=0.5,
    presence_penalty=0
    )
    print(response["choices"][0]["text"])
    tx1 = response["choices"][0]["text"]
    try:
        temp_list = []
        for i in tx1.split(":")[1:-1]:
            temp_list.append(i.split("\n")[0])
        temp_list.append(tx1.split(":")[-1])
    
        data = pd.read_excel(file)
        data["Score"] = temp_list
        data.to_excel("final_lead_score.xlsx")
        return data.to_html()
    except:
        return tx1
 
 
# Main Driver Function
if __name__ == '__main__':
    app.run(debug=True)