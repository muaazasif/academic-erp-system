import os
import sys
from app import app, db, Student

def migrate():
    with app.app_context():
        # Ensure is_active column exists
        try:
            db.session.execute(db.text("ALTER TABLE student ADD COLUMN IF NOT EXISTS is_active BOOLEAN DEFAULT TRUE;"))
            db.session.commit()
        except Exception:
            db.session.rollback()

        # Target lists of names for Bano Qabil 4.0 Summer
        bq_4_summer_names = [
            'Sidra Siddiqui', 'Humaira zeeshan', 'Ayman Nadeem', 'Unsa kamran', 'Laiba shahzad', 
            'Yusra kanwal', 'Mudassir', 'Muhammad Ishtiaque', 'Hasnain Mansoor', 'Sufyan Salahuddin', 
            'Syed Mohsib Shah', 'Muhammad Rayyan Sabir', 'Syeda Maheen Zehra Abidi', 'Muhammad Hassan', 
            'Maaz Ahmed Khan', 'Syed Zeeshan Ahmed', 'Abdur Rafay Kamran', 'Mawez Naqvi', 
            'Muhammad Usman Ali Khan Hannan', 'ALI SHEIKH', 'Muhammad AliWaiz', 'Taha Waqar', 
            'Neha Naz', 'Muhammad Farjad', 'MUHAMMAD TAYAYB', 'Muhammad Raheel', 'Masood Akhter', 
            'Rimsha Haroon', 'Muhammad Subhan Asif', 'Muhammad Ali', 'Kanwal jamil', 'Syeda Areeba', 
            'Musfirah Fatima', 'Mahad Ali Yousufzai', 'faisal ahmed'
        ]

        bq_4_fall_names = [
            'Wajahat', 'Muzammil Khalid', 'Danish Ghani', 'Maaz Ahmed', 'Zubair Khan', 'Muhammad Danial', 
            'saifullah', 'Rafay Badruddin', 'Rafia Javed', 'Shabeeb hassan', 'Alishba Ghazal', 
            'Sajjad Noor Ud Din', 'Shahzaib', 'Ibtihaj Ahmed', 'Muhammad shahzeb alam', 'babarkhanpak12@gmail.com', 
            'Zeeshan Aziz', 'Abdullah', 'Saadullah', 'Muhammad Furqan Khan', 'Asjal Azhar Khan', 
            'Narmeen Muskan', 'Muhammad Ali', 'Bisha Najam', 'Tooba Naushad', 'Hamza', 'Samina Perveen', 
            'Ifteha Hammad Zaheer', 'Waqar Khan', 'Hafiz Waiz Irshad', 'Muhammad Ahsan Baig', 'Muhammad Anas', 
            'Mujtaba hassan', 'Zainab', 'Ayesha Mahtab', 'Muhammad Zeeshan Israr', 'Hamdan.Sh', 
            'muhammad umair', 'Manal Asghar', 'Saifullah Khan', 'Ayaz Ahmed', 'Muhammad Hamza', 
            'Hassan', 'Muhammad Rizwan', 'Rashid shahbaz', 'M. Imran Sheikh', 'Maimoona Fahad', 
            'Mairaj Ahmed Khan', 'Wajahat Ali', 'Laiba Nazeer', 'salman'
        ]

        bq_5_summer_names = [
            'Mahek Naseem', 'Faisal Rehman Khan', 'Syed Muhammad Essa Shah', 'Azhar Khan Shehbaz', 
            'Hasnain fayyaz', 'Salman khan', 'Hammad', 'Muhammad Alyan', 'Muhammad shayan', 
            'NASEER IQBAL', 'Mehwish', 'Ribqa Asif', 'Muhammad Amir', 'Muhammad Aamir', 
            'elaaf', 'Javeria Ashraf', 'Raees Ahmed', 'Danial Ghani', 'Javeria Naqi'
        ]

        inactive_bq_5_fall_names = [
            'Muhammad Umar', 'Samama Ali', 'Maira Ali', 'Rabia Kashan', 'Abdul Rehman', 
            'Syed Nusrat Ali Hamidi', 'Syed Hashir Imam', 'Hifza mehtab', 'Syeda zainish fatima', 
            'Zainab', 'Maham shah', 'M.Rayyan Rahim', 'Hanzala khan', 'hayyan fahim', 'Muhammad Ahmed Cheema'
        ]

        count_4s = 0
        count_4f = 0
        count_5s = 0
        count_inactive = 0

        for s in Student.query.all():
            s_name_clean = s.name.strip().lower()
            
            # Check 4.0 Summer
            if any(n.strip().lower() == s_name_clean for n in bq_4_summer_names):
                s.batch = 'Bano Qabil 4.0 Summer'
                count_4s += 1
            # Check 4.0 Fall
            elif any(n.strip().lower() == s_name_clean for n in bq_4_fall_names):
                s.batch = 'Bano Qabil 4.0 Fall'
                count_4f += 1
            # Check 5.0 Summer
            elif any(n.strip().lower() == s_name_clean for n in bq_5_summer_names):
                s.batch = 'Bano Qabil 5.0 Summer'
                count_5s += 1
            # Check Inactive / Dropout 5.0 Fall
            elif any(n.strip().lower() == s_name_clean for n in inactive_bq_5_fall_names):
                s.batch = 'Bano Qabil 5.0 Fall'
                s.is_active = False
                count_inactive += 1

        db.session.commit()
        print(f"✅ Updated {count_4s} students to Bano Qabil 4.0 Summer")
        print(f"✅ Updated {count_4f} students to Bano Qabil 4.0 Fall")
        print(f"✅ Updated {count_5s} students to Bano Qabil 5.0 Summer")
        print(f"✅ Set {count_inactive} students in Bano Qabil 5.0 Fall as Inactive/Dropout")

if __name__ == "__main__":
    migrate()
