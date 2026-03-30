import json
import os
from schemas.models import parse_request
from core.generator import ExcelGeneratorService

with open('test_payload.json', 'r', encoding='utf-8') as f:
    data = json.load(f)
    
req = parse_request(data)
svc = ExcelGeneratorService(req)
output_io = svc.generate()

with open('output_test.xlsx', 'wb') as f:
    f.write(output_io.read())
    
print("Successfully generated output_test.xlsx")
