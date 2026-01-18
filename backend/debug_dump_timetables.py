from backend.app import app

c = app.test_client()
resp = c.get('/timetables')
arr = resp.get_json()
print('Total timetables returned:', len(arr))
filenames = [e.get('filename') for e in arr]
print('Unique filenames count:', len(set(filenames)))
print('\nSample filenames (first 100):')
for f in filenames[:100]:
    print('-', f)

# Show any filenames that contain '_classrooms'
classroom_files = [f for f in filenames if '_classrooms' in f]
print('\nFiles with "_classrooms" suffix:', len(classroom_files))
for f in classroom_files[:50]:
    print(' *', f)