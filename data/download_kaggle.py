"""
Download Employee Burnout dataset from GitHub mirrors of the Kaggle HackerEarth dataset.
"""
import requests
import os
import urllib3
import pandas as pd

urllib3.disable_warnings()

urls = [
    "https://raw.githubusercontent.com/AmanMalhotra123/HackerEarth_Employees_Burning_Out_Rank_09/main/train.csv",
    "https://raw.githubusercontent.com/AmanMalhotra123/HackerEarth_Employees_Burning_Out_Rank_09/master/train.csv",
    "https://raw.githubusercontent.com/Shivam-316/Employee-Burnout-Prediction/main/train.csv",
    "https://raw.githubusercontent.com/Shivam-316/Employee-Burnout-Prediction/master/train.csv",
    "https://raw.githubusercontent.com/ProjectByAbhi/Employee_Burnout_System/main/train.csv",
    "https://raw.githubusercontent.com/ProjectByAbhi/Employee_Burnout_System/master/train.csv",
    "https://raw.githubusercontent.com/DhairyaSethi/employee-burnout-analysis/main/train.csv",
    "https://raw.githubusercontent.com/DhairyaSethi/employee-burnout-analysis/master/train.csv",
    "https://raw.githubusercontent.com/SurajKumarRamagiri/Employee_burnout_analysis_and_prediction/main/train.csv",
    "https://raw.githubusercontent.com/SurajKumarRamagiri/Employee_burnout_analysis_and_prediction/master/train.csv",
]

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "kaggle_raw.csv")

for url in urls:
    try:
        r = requests.get(url, timeout=20, verify=False)
        repo = url.split("/")[4]
        print(f"[{r.status_code}] {repo} - {len(r.content)} bytes")
        if r.status_code == 200 and len(r.content) > 5000:
            with open(out, "wb") as f:
                f.write(r.content)
            print(f"SAVED to {out}")
            df = pd.read_csv(out)
            print(f"Rows: {len(df)}, Columns: {list(df.columns)}")
            print(f"Nulls:\n{df.isnull().sum()}")
            break
    except Exception as e:
        print(f"Error: {e}")
else:
    print("ALL MIRRORS FAILED - will need manual download from Kaggle")
