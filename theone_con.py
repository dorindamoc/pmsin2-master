from sqlalchemy import create_engine
import pandas as pd


theone_string  = 'postgresql://postgres:Renat0Dan1el@139.59.172.203:5432/THEONE'

def from_theone(q):
    engine = create_engine(theone_string)
    df = pd.read_sql(q, engine)
    return df