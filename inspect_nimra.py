from clean_sheets_sync import get_sheets_service
service, sheet_id = get_sheets_service()
for name in ['Final Marks', 'Final Results']:
    try:
        res = service.spreadsheets().values().get(spreadsheetId=sheet_id, range=f"{name}!A:Z").execute()
        rows = res.get('values', [])
        headers = rows[0]
        print(f"--- Sheet: {name} ---")
        for r in rows[1:]:
            row_str = " ".join([str(c) for c in r])
            if '557727' in row_str or 'Nimra' in row_str:
                print(dict(zip(headers, r)))
    except Exception as e:
        print(f"Error {name}: {e}")
