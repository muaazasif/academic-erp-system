from google_classroom_attendance_bot import get_classroom_service

service = get_classroom_service()
if service:
    results = service.courses().list(pageSize=10).execute()
    courses = results.get('courses', [])
    print(f"Found {len(courses)} courses:")
    for c in courses:
        print(f"Name: {c.get('name')} | ID: {c.get('id')} | Section: {c.get('section')}")
else:
    print("Service not available or not authenticated")
