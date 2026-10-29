from google_classroom_attendance_bot import get_classroom_service

def list_my_courses():
    service = get_classroom_service()
    if not service:
        print("❌ Not authenticated with Google Classroom. Please authenticate first via the admin panel.")
        return
    
    try:
        results = service.courses().list(pageSize=10).execute()
        courses = results.get('courses', [])
        
        if not courses:
            print("⚠️ No courses found for this Google account.")
        else:
            print("✅ Found Google Classroom courses:")
            for course in courses:
                print(f"  - Name: {course.get('name')}")
                print(f"    ID: {course.get('id')}")
                print(f"    Section: {course.get('section', 'N/A')}")
                print("-" * 40)
    except Exception as e:
        print(f"❌ Error listing courses: {e}")

if __name__ == '__main__':
    list_my_courses()
