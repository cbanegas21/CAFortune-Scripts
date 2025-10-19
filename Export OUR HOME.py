import os
import pandas as pd
import pyodbc

# —————————————————————————————————————————
#  YOUR AZURE SQL CREDENTIALS
# —————————————————————————————————————————
SERVER   = 'ca-data-server.database.windows.net'
DATABASE = 'CAFortuneDatabase'
USERNAME = 'sqladmin'
PASSWORD = 'Maxine2021.'
DRIVER   = '{ODBC Driver 17 for SQL Server}'

# —————————————————————————————————————————
#  OUTPUT FOLDER
# —————————————————————————————————————————
OUT_DIR = './exports'
os.makedirs(OUT_DIR, exist_ok=True)

# —————————————————————————————————————————
#  TABLES TO DUMP AS CSV
#  (all from your screenshots)
# —————————————————————————————————————————
tables = [
    # Food Should Taste Good
    "food_should_taste_good_aci_jewel_osco",
    "food_should_taste_good_fresh_thyme",
    "food_should_taste_good_hannaford",
    "food_should_taste_good_natural_grocers",
    "food_should_taste_good_our_home_independents",
    "food_should_taste_good_shaws",
    "food_should_taste_good_stopshop",
    "food_should_taste_good_upc_audit",
    "food_should_taste_good_upc_audit_whole_foods",
    "food_should_taste_good_wakefern",
    "food_should_taste_good_wegmans",
    "food_should_taste_good_whole_foods_market",

    # From The Ground Up
    "from_the_ground_up_aci_jewel_osco",
    "from_the_ground_up_fresh_thyme",
    "from_the_ground_up_hannaford",
    "from_the_ground_up_heb",
    "from_the_ground_up_natural_grocers",
    "from_the_ground_up_publix",
    "from_the_ground_up_the_fresh_market",
    "from_the_ground_up_wakefern",
    "from_the_ground_up_wegmans",
    "from_the_ground_up_whole_foods_market",

    # Good Health
    "good_health_fresh_thyme",
    "good_health_hannaford",
    "good_health_harris_teeter",
    "good_health_natural_grocers",
    "good_health_publix",
    "good_health_schnucks",
    "good_health_shaws",
    "good_health_stopshop",
    "good_health_the_fresh_market",
    "good_health_wakefern",
    "good_health_wegmans",

    # ParmCrisps
    "parmcrisps_aci_jewel_osco",
    "parmcrisps_fresh_thyme",
    "parmcrisps_hain",
    "parmcrisps_hannaford",
    "parmcrisps_harris_teeter",
    "parmcrisps_natural_grocers",
    "parmcrisps_publix",
    "parmcrisps_raleys",
    "parmcrisps_shaws",
    "parmcrisps_stopshop",
    "parmcrisps_the_fresh_market",
    "parmcrisps_wakefern",
    "parmcrisps_wegmans",
    "parmcrisps_whole_foods_market",

    # Pop Secret
    "pop_secret_aci_jewel_osco",
    "pop_secret_fresh_thyme",
    "pop_secret_harris_teeter",
    "pop_secret_heb",
    "pop_secret_heb_our_home",
    "pop_secret_publix",
    "pop_secret_raleys",
    "pop_secret_shaws",
    "pop_secret_wakefern",
    "pop_secret_wegmans",
    "pop_secret_wegmans_our_home",

    # Popchips
    "popchips_fresh_thyme",
    "popchips_hannaford",
    "popchips_harris_teeter",
    "popchips_heb",
    "popchips_natural_grocers",
    "popchips_publix",
    "popchips_raleys",
    "popchips_shaws",
    "popchips_stopshop",
    "popchips_wakefern",
    "popchips_wegmans",

    # RW Garcia
    "rw_garcia_central_market",
    "rw_garcia_fresh_thyme",
    "rw_garcia_hannaford",
    "rw_garcia_harris_teeter",
    "rw_garcia_natural_grocers",
    "rw_garcia_our_home_wf",
    "rw_garcia_stopshop",
    "rw_garcia_wakefern",
    "rw_garcia_wegmans",
    "rw_garcia_whole_foods_market",

    # Sonoma Creamery
    "sonoma_creamery_aci_jewel_osco",
    "sonoma_creamery_central_market",
    "sonoma_creamery_harris_teeter",
    "sonoma_creamery_raleys",
    "sonoma_creamery_schnucks",
    "sonoma_creamery_the_fresh_market",
    "sonoma_creamery_wakefern",

    # You Need This (YNT)
    "you_need_this_fresh_thyme",
    "you_need_this_heb",
    "you_need_this_wakefern",
]

# —————————————————————————————————————————
#  CONNECT & EXPORT LOOP
# —————————————————————————————————————————
conn_str = (
    f"DRIVER={DRIVER};"
    f"SERVER={SERVER};"
    f"DATABASE={DATABASE};"
    f"UID={USERNAME};"
    f"PWD={PASSWORD}"
)
cnxn = pyodbc.connect(conn_str)

for tbl in tables:
    fq_table = f"dbo.[{tbl}]"
    csv_path = os.path.join(OUT_DIR, f"{tbl}.csv")
    print(f"Exporting {fq_table} -> {csv_path} ...", end=" ")
    try:
        df = pd.read_sql(f"SELECT * FROM {fq_table}", cnxn)
        df.to_csv(csv_path, index=False)
        print("OK")
    except Exception as e:
        print("FAILED:", e)

cnxn.close()
print("All tables exported!")
