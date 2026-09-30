import os
import json
import sys
import time
from datetime import datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from werkzeug.security import generate_password_hash

# Add the app directory to the path to import models
sys.path.insert(0, '.')
from app import app, db, Student

# Google Sheet Configuration
SPREADSHEET_ID = os.environ.get('GOOGLE_SHEET_ID', '1N23HvM_BvBEKDVi-q1m76ZozEmk1MSbPMitW6rltK5c')
SHEET_NAME = 'username'  # User requested sheet name 'username'
# Fallback sheet names if 'username' is not found
FALLBACK_SHEET_NAMES = ['users', 'Form Responses 1', 'Sheet1']

LOCK_FILE = 'instance/sync_users.lock'

from clean_sheets_sync import get_sheets_service

def is_process_alive(pid):
    if not pid:
        return False
    try:
        pid = int(pid)
        if pid <= 0:
            return False
        os.kill(pid, 0)
        return True
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    except Exception:
        return False

def sync_users_from_sheet():
    os.makedirs('instance', exist_ok=True)
    
    # Safe stale-lock recovery & concurrency check
    if os.path.exists(LOCK_FILE):
        try:
            with open(LOCK_FILE, 'r') as f:
                content = f.read().strip()
                lock_data = {}
                try:
                    lock_data = json.loads(content)
                except json.JSONDecodeError:
                    lock_data = {'pid': content, 'timestamp': os.path.getmtime(LOCK_FILE)}
                
                lock_time = float(lock_data.get('timestamp', os.path.getmtime(LOCK_FILE)))
                lock_pid = lock_data.get('pid')
                age = time.time() - lock_time
                
                # Check if lock is stale (older than 10 minutes OR recorded process is dead)
                is_stale = (age > 600) or (lock_pid and not is_process_alive(lock_pid))
                
                if not is_stale:
                    print(f"⚠️ Sync already in progress (locked at {datetime.fromtimestamp(lock_time)}). Skipping.")
                    return
                else:
                    print(f"⚠️ Stale or orphaned lock file found (age: {int(age)}s, PID: {lock_pid}), removing it.")
                    os.remove(LOCK_FILE)
        except Exception as e:
            print(f"⚠️ Error reading lock file: {e}, removing it.")
            try:
                os.remove(LOCK_FILE)
            except:
                pass

    # Acquire lock atomically
    try:
        fd = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        with os.fdopen(fd, 'w') as f:
            lock_info = {
                'pid': os.getpid(),
                'timestamp': time.time(),
                'datetime': datetime.now().isoformat()
            }
            json.dump(lock_info, f)
    except FileExistsError:
        print("⚠️ Sync already in progress (lock file exists). Skipping.")
        return
    except Exception as e:
        print(f"⚠️ Failed to create lock file: {e}")
        return

    try:
        print(f"🚀 Starting User Sync from Google Sheet: {SPREADSHEET_ID}")
        
        # Use the existing service getter from clean_sheets_sync
        service, _ = get_sheets_service()
        if not service:
            print("❌ Could not initialize Google Sheets service. Check environment variables.")
            return

        # First, try to find which sheet exists
        spreadsheet = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        sheets = [s['properties']['title'] for s in spreadsheet.get('sheets', [])]
        
        target_sheet = None
        if SHEET_NAME in sheets:
            target_sheet = SHEET_NAME
        else:
            for fallback in FALLBACK_SHEET_NAMES:
                if fallback in sheets:
                    target_sheet = fallback
                    break
        
        if not target_sheet:
            print(f"❌ Could not find target sheet '{SHEET_NAME}' or any fallbacks in {sheets}")
            return

        print(f"Using sheet: {target_sheet}")
        
        # Fetch data from the sheet
        range_name = f"'{target_sheet}'!A:Z"
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        if not values:
            print("No data found in the sheet.")
            return

        # Header is values[0], data starts from values[1]
        headers = values[0]
        print(f"Headers found: {headers}")
        
        # Dynamically find column indices
        student_id_col = -1
        name_col = -1
        
        # Keywords to look for
        student_id_keywords = ['studentid', 'roll number', 'student id', 'roll_number', 'user id', 'username']
        name_keywords = ['login name', 'name', 'full name', 'student name', 'login_name']
        
        for i, header in enumerate(headers):
            header_lower = str(header).lower().strip()
            if student_id_col == -1 and any(k in header_lower for k in student_id_keywords):
                student_id_col = i
            elif name_col == -1 and any(k in header_lower for k in name_keywords):
                name_col = i

        # If not found, use defaults
        if student_id_col == -1:
            student_id_col = 0 # Default to first column
            print(f"⚠️ StudentID column not found by header, using column index {student_id_col}")
        if name_col == -1:
            name_col = 1 if len(headers) > 1 else 0
            print(f"⚠️ Name column not found by header, using column index {name_col}")

        print(f"Using columns: StudentID at {student_id_col}, Name at {name_col}")

        new_users_count = 0
        existing_users_count = 0
        
        with app.app_context():
            db.create_all()
            # Get all existing student IDs in one go to avoid autoflush issues
            existing_student_ids = {s.student_id for s in Student.query.with_entities(Student.student_id).all()}
            processed_in_this_run = set()
            
            with db.session.no_autoflush:
                for i, row in enumerate(values[1:], start=2):
                    if len(row) <= max(student_id_col, name_col):
                        # Pad row if it's shorter than expected columns
                        row.extend([''] * (max(student_id_col, name_col) - len(row) + 1))
                    
                    student_id = str(row[student_id_col]).strip()
                    if not student_id or student_id.lower() in ['studentid', 'roll number', 'student id', 'roll_number', 'user id', 'username']:
                        continue
                    
                    # Skip if we already processed this ID in this run
                    if student_id in processed_in_this_run:
                        continue
                    processed_in_this_run.add(student_id)
                    
                    name = str(row[name_col]).strip() if name_col < len(row) else student_id
                    if not name:
                        name = student_id
                    
                    # Check if student already exists in our pre-fetched set
                    if student_id not in existing_student_ids:
                        print(f"➕ Creating new user: {student_id} ({name})")
                        try:
                            new_student = Student(
                                student_id=student_id,
                                name=name
                            )
                            # Username and Password are the same (student_id)
                            new_student.set_password(student_id)
                            db.session.add(new_student)
                            new_users_count += 1
                        except Exception as inner_e:
                            print(f"⚠️ Failed to add student {student_id}: {inner_e}")
                    else:
                        existing_users_count += 1
            
            if new_users_count > 0:
                try:
                    db.session.commit()
                    print(f"✅ Successfully created {new_users_count} new users.")
                except Exception as commit_e:
                    db.session.rollback()
                    print(f"❌ Failed to commit new users: {commit_e}")
            else:
                print("ℹ️ No new users to create.")
            
            print(f"📊 Total processed: {len(processed_in_this_run)}")
            print(f"📊 Existing users: {existing_users_count}")

    except Exception as e:
        print(f"❌ Error during sync: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # Always remove lock file
        if os.path.exists(LOCK_FILE):
            try:
                os.remove(LOCK_FILE)
            except Exception as e:
                print(f"⚠️ Error removing lock file: {e}")

if __name__ == "__main__":
    sync_users_from_sheet()

if __name__ == "__main__":
    sync_users_from_sheet()
