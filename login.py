from pymongo import MongoClient
from urllib.parse import quote_plus

username = "Sachin"
password = "NBsjkdIU#@Y20"
host = "192.168.1.144"

uri = (
    f"mongodb://{quote_plus(username)}:{quote_plus(password)}@{host}:27017/CRM_Main?authSource=admin"
)

client = MongoClient(uri, serverSelectionTimeoutMS=5000)

try:
    client.admin.command("ping")
    print("MongoDB connected successfully!")

    db = client["Final_Code"]

     # Access CRM_Main collection
    crm_collection = db["CRM_Main"]

    print("===== CRM_Main =====")
    for document in crm_collection.find():
        print(document)

    # Access Companies collection
    companies_collection = db["Companies"]

    print("\n===== Companies =====")
    for document in companies_collection.find():
        print(document)

except Exception as e:
    print("MongoDB connection failed:", e)

finally:
    client.close()

