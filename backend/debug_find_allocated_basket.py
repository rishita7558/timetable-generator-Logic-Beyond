import sys, os
sys.path.insert(0, os.path.join(os.getcwd(), 'backend'))
from app import app

c = app.test_client()
r = c.get('/timetables')
d = r.get_json()

# Find a timetable with non-empty basket_course_allocations
for t in d:
    bca = t.get('basket_course_allocations')
    if isinstance(bca, dict):
        for basket, mapping in bca.items():
            if mapping and any(v is not None for v in mapping.values()):
                print('Found timetable with allocations:', t.get('filename'))
                print('basket_course_allocations sample:', {basket: mapping})
                print('HTML sample snippet:', t.get('html')[:400])
                raise SystemExit(0)

print('No timetable with non-empty basket course allocations found in current dataset.')
