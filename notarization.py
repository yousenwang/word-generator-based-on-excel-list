from pathlib import Path
import math
import os
import pandas as pd  # pip install pandas openpyxl
from docxtpl import DocxTemplate  # pip install docxtpl


def get_contract_content(df):
    content = {}
    for i, record in enumerate(df.to_dict(orient="records")):
        for key in list(record.keys()):
            record[key+str(i)] = record.pop(key)
        content.update(record)
    return content


def get_skip_rows(excel_path, sheet_name, case_number):
    """
    只讀取案號，節省暫存空間。
    在此找出非本案號的行號碼們，返回進行篩選。
    """
    df = pd.read_excel(
        excel_path,
        sheet_name=sheet_name,
        usecols=["案號"]
        )
    return [x+1 for x in df[df["案號"] != case_number].index]

def contract(file_name, case_number):
    base_dir = Path(__file__).parent
    input_folder = "公證樣本輸入"
    word_template_path = base_dir / input_folder / f"{file_name}template.docx"
    excel_path = base_dir / input_folder /"公證輸入.xlsx"
    output_dir = base_dir / "公證輸出"
    # Create output folder for the word documents
    output_dir.mkdir(exist_ok=True)

    # Convert Excel sheet to pandas dataframe
    row_numbers = get_skip_rows(
        excel_path,
        "roles",
        case_number
    )

    df = pd.read_excel(
        excel_path,
        sheet_name="roles",
        header=0,
        skiprows=row_numbers
    )
    ##print(df)
    df.fillna('', inplace=True)
    # Iterate over each row in df and render word document
    contract_content = {"people" : list(df.to_dict(orient="index").values())}
    print(contract_content)
    ##contract_content = get_contract_content(df)
    
    ##print(case_content)
    # 寫入word document
    doc = DocxTemplate(word_template_path)
    doc.render(contract_content)
    output_path = output_dir / f"{case_number}_{file_name}.docx"##f"{record['統一編號/身分證字號']}-租賃公證.docx"
    doc.save(output_path)

def request(case_number):
    base_dir = Path(__file__).parent
    input_folder = "公證樣本輸入"
    excel_path = base_dir / input_folder /"公證輸入.xlsx"
    output_dir = base_dir / "公證輸出"
    # Create output folder for the word documents
    output_dir.mkdir(exist_ok=True)

    # Convert Excel sheet to pandas dataframe
    row_numbers = get_skip_rows(
        excel_path,
        "roles",
        case_number
    )

    df = pd.read_excel(
        excel_path,
        sheet_name="roles",
        header=0,
        skiprows=row_numbers
    )
    df.fillna('', inplace=True)
    
    # 寫入word document

    row_numbers = get_skip_rows(
        excel_path,
        "cases",
        case_number
    )
    case_content = pd.read_excel(
        excel_path,
        sheet_name="cases",
        header=0,
        skiprows=row_numbers
    ).to_dict(orient="records")[0]

    def render_output(file_name, context):
        word_template_path = base_dir / input_folder / f"{file_name}template.docx"
        doc = DocxTemplate(word_template_path)
        doc.render(context)
        output_path = output_dir / f"{case_number}_{file_name}.docx"##f"{record['統一編號/身分證字號']}-租賃公證.docx"
        doc.save(output_path)
        
        print(f"{output_path} is rendered and saved.")
    
    request_context = {
        'applicants': df[df["角色類型"] == "請求人"].to_dict(orient="records"),
        'agents': df[df["角色類型"] == "代理人"].to_dict(orient="records"),
        'third_parties': df[df["角色類型"] == "第三人"].to_dict(orient="records")
    }
    request_context.update(case_content)
    render_output("公證請求書", request_context)
    render_output("公證卷宗", request_context)

    render_output("公證例稿", request_context)
    render_output("收據", case_content)



    contract_content = {"people" : list(df.to_dict(orient="index").values())}
    render_output("土地租賃契約", contract_content)
    ##content.update(pd.read_excel(excel_path, sheet_name="details").to_dict(orient="records")[0])
    
if __name__ == "__main__":
    print("notarization")
    print("Please enter the case number: ")
    case_number = input()
    
    request(int(case_number))


    # print("You entered: " + case_number)
    # contract("土地租賃契約", int(case_number))
    
    ##contract("土地租賃契約", 10002)



