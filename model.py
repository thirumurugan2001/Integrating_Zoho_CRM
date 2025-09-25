from pydantic import BaseModel

class Leads(BaseModel):
    leads: list

from pydantic import BaseModel

class FilePath(BaseModel):
    file_path: str
