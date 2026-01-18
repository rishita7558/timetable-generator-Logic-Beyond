import os
import shutil
import pandas as pd
from backend.app import INPUT_DIR, load_all_data


def _backup_and_write_faculty_csv(tmp_csv_content):
    orig_path = os.path.join(INPUT_DIR, 'faculty_availability.csv')
    backup_path = os.path.join(INPUT_DIR, 'faculty_availability.csv.bak')
    if os.path.exists(orig_path):
        shutil.copy(orig_path, backup_path)
    with open(orig_path, 'w', encoding='utf-8') as f:
        f.write(tmp_csv_content)
    return backup_path if os.path.exists(backup_path) else None


def _restore_faculty_csv(backup_path):
    path = os.path.join(INPUT_DIR, 'faculty_availability.csv')
    if backup_path and os.path.exists(backup_path):
        shutil.move(backup_path, path)
    else:
        if os.path.exists(path):
            os.remove(path)


def test_single_column_faculty_availability_normalized(tmp_path):
    csv_content = "FACULTY NAME\nAlice\nBob\n"
    backup = _backup_and_write_faculty_csv(csv_content)

    dfs = load_all_data(force_reload=True)
    assert dfs is not None
    assert 'faculty_availability' in dfs

    fa = dfs['faculty_availability']
    # Columns should include normalized fields
    # Column names should be normalized to expected exact names
    assert 'Faculty Name' in fa.columns
    assert 'Available Days' in fa.columns
    assert 'Unavailable Time Slots' in fa.columns

    # Defaults should be set
    assert fa['Available Days'].iloc[0] == 'Mon,Tue,Wed,Thu,Fri'
    assert str(fa['Unavailable Time Slots'].iloc[0]) == ''

    _restore_faculty_csv(backup)
