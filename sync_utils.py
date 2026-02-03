import os
import json
from datetime import datetime, time
import threading

# Global lock for thread-safe file operations
file_lock = threading.Lock()

def store_failed_sync(data_type, data):
    """
    Store failed Google Sheets sync attempts to a local file for later processing
    """
    failed_syncs_file = 'failed_syncs.json'

    # Convert any complex objects to JSON serializable format
    serializable_data = {}
    for key, value in data.items():
        if isinstance(value, (datetime, time)):
            # Convert datetime objects to string
            serializable_data[key] = str(value)
        elif hasattr(value, '__dict__'):  # Object with attributes
            # Convert object to dictionary representation
            if hasattr(value, '__tablename__'):  # SQLAlchemy model
                # For SQLAlchemy models, extract basic attributes
                serializable_data[key] = str(value)
            else:
                # For other objects, try to convert to dict
                try:
                    serializable_data[key] = vars(value)
                except:
                    serializable_data[key] = str(value)
        elif isinstance(value, (list, dict)):
            # Handle lists and dictionaries by converting to string if they contain non-serializable items
            try:
                # Test if it's JSON serializable
                import json
                json.dumps(value)
                serializable_data[key] = value
            except (TypeError, ValueError):
                # If not serializable, convert to string representation
                serializable_data[key] = str(value)
        else:
            # Other values should be fine
            serializable_data[key] = value

    with file_lock:
        # Load existing failed syncs
        failed_syncs = []
        if os.path.exists(failed_syncs_file):
            try:
                with open(failed_syncs_file, 'r') as f:
                    failed_syncs = json.load(f)
            except:
                failed_syncs = []

        # Add new failed sync
        failed_syncs.append({
            'type': data_type,
            'data': serializable_data,
            'timestamp': datetime.now().isoformat(),
            'status': 'pending_retry'
        })

        # Save back to file
        with open(failed_syncs_file, 'w') as f:
            json.dump(failed_syncs, f, indent=2)

        print(f"Stored failed {data_type} sync for later retry")


def retry_failed_syncs():
    """
    Attempt to retry previously failed Google Sheets syncs
    """
    failed_syncs_file = 'failed_syncs.json'

    if not os.path.exists(failed_syncs_file):
        return 0

    with file_lock:
        try:
            with open(failed_syncs_file, 'r') as f:
                failed_syncs = json.load(f)
        except:
            return 0

    # Filter out completed syncs
    pending_syncs = [sync for sync in failed_syncs if sync.get('status') == 'pending_retry']

    if not pending_syncs:
        return 0

    # Import the sync functions
    from app import add_attendance_to_sheet, add_assignment_submission_to_sheet, \
                 add_quiz_submission_to_sheet, add_detailed_quiz_answers_to_sheet, \
                 add_midterm_grade_to_sheet

    successful_retries = 0

    for sync in pending_syncs:
        try:
            sync_type = sync['type']
            sync_data = sync['data']

            success = False
            if sync_type == 'attendance':
                success = add_attendance_to_sheet(
                    student_id=sync_data['student_id'],
                    name=sync_data['name'],
                    date=sync_data['date'],
                    check_in=sync_data['check_in'],
                    check_out=sync_data['check_out'],
                    status=sync_data['status']
                )
            elif sync_type == 'assignment':
                success = add_assignment_submission_to_sheet(
                    student_id=sync_data['student_id'],
                    name=sync_data['name'],
                    assignment_title=sync_data['assignment_title'],
                    submission_url=sync_data['submission_url'],
                    submitted_at=sync_data['submitted_at'],
                    grade=sync_data.get('grade')
                )
            elif sync_type == 'quiz':
                success = add_quiz_submission_to_sheet(
                    student_id=sync_data['student_id'],
                    name=sync_data['name'],
                    quiz_title=sync_data['quiz_title'],
                    score=sync_data['score'],
                    total_questions=sync_data['total_questions'],
                    submitted_at=sync_data['submitted_at']
                )
            elif sync_type == 'detailed_quiz':
                # For detailed quiz data, we need to reconstruct the objects properly
                # Since questions_data contains SQLAlchemy objects that can't be serialized,
                # we'll skip retrying this type for now to prevent errors
                print(f"Skipping retry for detailed_quiz sync due to complex object serialization issues")
                success = False
                # Mark as successful to prevent repeated retries of problematic data
                sync['status'] = 'skipped_complex'
                continue
            elif sync_type == 'midterm_grade':
                success = add_midterm_grade_to_sheet(
                    student_id=sync_data['student_id'],
                    name=sync_data['name'],
                    midterm_title=sync_data['midterm_title'],
                    grade=sync_data['grade'],
                    graded_at=sync_data['graded_at']
                )

            if success:
                # Mark as successful
                sync['status'] = 'successful'
                successful_retries += 1
                print(f"Successfully retried {sync_type} sync")
            else:
                print(f"Retry failed for {sync_type} sync")
                # Increase retry count to eventually give up on problematic entries
                if 'retry_count' not in sync:
                    sync['retry_count'] = 1
                else:
                    sync['retry_count'] += 1

                # If retry count exceeds 10, mark as failed permanently
                if sync['retry_count'] >= 10:
                    sync['status'] = 'failed_permanently'
                    print(f"Marking {sync_type} sync as permanently failed after 10 attempts")

        except Exception as e:
            print(f"Error during retry of {sync.get('type', 'unknown')} sync: {e}")
            # Log the error but continue with other syncs
            import traceback
            traceback.print_exc()

    # Save updated status
    with open(failed_syncs_file, 'w') as f:
        json.dump(failed_syncs, f, indent=2)

    return successful_retries


def cleanup_old_failed_syncs(days_to_keep=30):
    """
    Remove failed sync attempts older than specified days
    """
    failed_syncs_file = 'failed_syncs.json'
    
    if not os.path.exists(failed_syncs_file):
        return 0
    
    with file_lock:
        try:
            with open(failed_syncs_file, 'r') as f:
                failed_syncs = json.load(f)
        except:
            return 0
    
    import datetime as dt
    
    cutoff_date = datetime.now() - dt.timedelta(days=days_to_keep)
    remaining_syncs = []
    removed_count = 0
    
    for sync in failed_syncs:
        sync_time = dt.datetime.fromisoformat(sync['timestamp'])
        if sync_time > cutoff_date:
            remaining_syncs.append(sync)
        else:
            removed_count += 1
    
    if removed_count > 0:
        with open(failed_syncs_file, 'w') as f:
            json.dump(remaining_syncs, f, indent=2)
        print(f"Removed {removed_count} old failed sync attempts")
    
    return removed_count


def background_sync_worker():
    """
    Background worker to periodically retry failed syncs
    """
    while True:
        try:
            successful_retries = retry_failed_syncs()
            if successful_retries > 0:
                print(f"Background sync worker: Successfully retried {successful_retries} syncs")
            
            # Clean up old failed syncs occasionally
            if int(time.time()) % 3600 < 60:  # Every hour check for cleanup
                cleanup_old_failed_syncs()
            
            # Wait 5 minutes before next check
            time.sleep(300)
        except Exception as e:
            print(f"Error in background sync worker: {e}")
            time.sleep(60)  # Wait 1 minute before retrying


def start_background_sync():
    """
    Start the background sync worker in a separate thread
    """
    sync_thread = threading.Thread(target=background_sync_worker, daemon=True)
    sync_thread.start()
    print("Started background sync worker thread")