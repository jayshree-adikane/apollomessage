import os
from dotenv import load_dotenv

# Secrets and credentials come from the local .env file (see .env.example).
load_dotenv()

import json
from pymongo import MongoClient
from bson import ObjectId

# ============================================================
# MongoDB
# ============================================================

client = MongoClient(
    os.getenv("MONGODB_URI", "")
)

db = client["Apollo_Email"]

collection = db["prospects"]
blacklist_collection = db["Blacklist"]


# ============================================================
# FIND PROFILES WHERE email_sent IS NULL
# ============================================================

query = {
    "email_sent": None
}

docs = list(collection.find(query))

print(f"Found {len(docs)} documents to move.")


# ============================================================
# MOVE TO BLACKLIST
# ============================================================

if docs:

    try:
        # Insert documents into Blacklist
        result = blacklist_collection.insert_many(
            docs,
            ordered=False
        )

        print(f"Inserted {len(result.inserted_ids)} documents into Blacklist.")

        # Delete ONLY the documents successfully inserted
        collection.delete_many({
            "_id": {
                "$in": result.inserted_ids
            }
        })

        print("Successfully deleted moved documents from prospects.")

    except Exception as e:
        print("Error while moving documents:")
        print(e)

else:
    print("No documents found where email_sent is null.")


# ============================================================
# CLOSE CONNECTION
# ============================================================

client.close()