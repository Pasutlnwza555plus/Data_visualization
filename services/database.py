import pandas as pd
import pymongo
import os
from dotenv import load_dotenv
from pathlib import Path

class Database:
    def __init__(self):
        env_path = Path(__file__).resolve().parent.parent / ".env"
        load_dotenv(dotenv_path=env_path)

        mongo_host_uri = os.getenv("mongo_host_uri")
        if not mongo_host_uri:
            raise ValueError("Environment variable 'mongo_host_uri' not set")

        self.mongo_client = pymongo.MongoClient(mongo_host_uri)
        self.database = self.mongo_client["ZTE"]
        self.reference_collection = self.database["ReferenceSheet"]

    def get_reference_sheet(self, query=None):
        links = list(self.reference_collection.find(query if query else {}))
        
        return pd.DataFrame(links).drop("_id", axis=1)