from backend.app import app
client = app.test_client()
resp = client.get('/timetables')
data = resp.get_json()
for t in data:
    if t.get('filename') == 'sem3_ECE_forced_conflict_classrooms.xlsx':
        print('FOUND', t.get('section'))
        for d in t.get('classroom_details', []):
            print(d)
