# SQL Assignment Solutions & Grading Guide

This guide contains the tasks, correct queries, and a troubleshooting guide explaining why a student's query might be logically correct but marked as **incorrect** by the automated grading system.

---

## 1. SQL Advanced Level: Complex Queries (Solution)

These are the exact 10 questions and their expected queries as defined in the system for **SQL Advanced Level: Complex Queries**:

### Task 1
* **Task:** Create a VIEW 'DetailedReport' joining Students, Enrollments, and Courses with all details.
* **Expected Query:**
  ```sql
  SELECT Students.name, Students.city, Courses.title, Courses.fee, Enrollments.enrollment_date 
  FROM Students 
  JOIN Enrollments ON Students.id = Enrollments.student_id 
  JOIN Courses ON Enrollments.course_id = Courses.id
  ```

### Task 2
* **Task:** Find students who are enrolled in more than 1 course.
* **Expected Query:**
  ```sql
  SELECT name 
  FROM Students 
  WHERE id IN (SELECT student_id FROM Enrollments GROUP BY student_id HAVING COUNT(*) > 1)
  ```

### Task 3
* **Task:** Use a CTE (WITH clause) to list students from 'Karachi' and their total course fees.
* **Expected Query:**
  ```sql
  WITH StudentFees AS (
      SELECT student_id, SUM(fee) as total 
      FROM Enrollments 
      JOIN Courses ON Enrollments.course_id = Courses.id 
      GROUP BY student_id
  ) 
  SELECT name, total 
  FROM Students 
  JOIN StudentFees ON Students.id = StudentFees.student_id 
  WHERE city = 'Karachi'
  ```

### Task 4
* **Task:** Find the top 2 highest paying students and their names.
* **Expected Query:**
  ```sql
  SELECT name, SUM(fee) 
  FROM Students 
  JOIN Enrollments ON Students.id = Enrollments.student_id 
  JOIN Courses ON Enrollments.course_id = Courses.id 
  GROUP BY name 
  ORDER BY SUM(fee) DESC 
  LIMIT 2
  ```

### Task 5
* **Task:** Find courses that have no enrollments.
* **Expected Query:**
  ```sql
  SELECT title 
  FROM Courses 
  WHERE id NOT IN (SELECT course_id FROM Enrollments)
  ```

### Task 6
* **Task:** Show student names and a column 'Status' which is 'Karachi Resident' if they live in Karachi, otherwise 'Other'.
* **Expected Query:**
  ```sql
  SELECT name, CASE WHEN city = 'Karachi' THEN 'Karachi Resident' ELSE 'Other' END as Status 
  FROM Students
  ```

### Task 7
* **Task:** Find the student who enrolled first in any course.
* **Expected Query:**
  ```sql
  SELECT name 
  FROM Students 
  JOIN Enrollments ON Students.id = Enrollments.student_id 
  ORDER BY enrollment_date ASC 
  LIMIT 1
  ```

### Task 8
* **Task:** Get a list of all cities and the total revenue from each city.
* **Expected Query:**
  ```sql
  SELECT Students.city, SUM(Courses.fee) 
  FROM Students 
  JOIN Enrollments ON Students.id = Enrollments.student_id 
  JOIN Courses ON Enrollments.course_id = Courses.id 
  GROUP BY Students.city
  ```

### Task 9
* **Task:** Find students who live in the same city as 'Sara Ahmed'.
* **Expected Query:**
  ```sql
  SELECT name 
  FROM Students 
  WHERE city = (SELECT city FROM Students WHERE name = 'Sara Ahmed') 
    AND name != 'Sara Ahmed'
  ```

### Task 10
* **Task:** Calculate the percentage of total students that live in each city.
* **Expected Query:**
  ```sql
  SELECT city, COUNT(*)*100.0 / (SELECT COUNT(*) FROM Students) 
  FROM Students 
  GROUP BY city
  ```

---

## 2. Why does the grader mark correct answers as WRONG?

The grading engine (`sql_grader.py`) compares the student's result table (as a Pandas DataFrame) with the expected result table. Even if the student's SQL is logically 100% correct and retrieves the right data, it might be marked incorrect due to the following reasons:

### A. Column Names & Aliases Mismatch (Most Common)
If the student renames columns using `AS` or fails to use the expected aggregate function name, the column names will not match.
* **Example (Task 4):**
  * Expected query uses: `SELECT name, SUM(fee) ...`
  * Student query uses: `SELECT name, SUM(fee) AS [Total Paid] ...`
  * **Result:** **Incorrect** because the column is named `Total Paid` instead of `SUM(fee)`.
* **Example (Task 6):**
  * Expected query uses: `... END as Status ...` (case-sensitive)
  * Student query uses: `... END as status ...` (lowercase)
  * **Result:** **Incorrect** due to case-sensitive column headers.

### B. Column Selection Order
The order in which columns are selected must be exactly the same as the expected query.
* **Example (Task 8):**
  * Expected: `SELECT Students.city, SUM(Courses.fee) ...`
  * Student: `SELECT SUM(Courses.fee), Students.city ...`
  * **Result:** **Incorrect** because the columns are swapped.

### C. Sorting / Ordering Difference
For tasks where sorting is required (using `ORDER BY`), the grader expects the exact row sequence. If the student sorts differently, or does not include the exact `ORDER BY` clause, it will fail.
* **Example (Task 4 & Task 7):**
  * If the student uses `ORDER BY SUM(fee)` instead of `ORDER BY SUM(fee) DESC`, or if sorting columns are different, it is marked **Incorrect**.

### D. Exact String Matches (Case and Whitespace Sensitivity)
If a task compares text (like 'Karachi'), the query must match the casing and string representation perfectly.
* **Example (Task 6):**
  * Expected: `'Karachi Resident'`
  * Student: `'karachi resident'` (lowercase) or `'Karachi resident'`
  * **Result:** **Incorrect** because the string values differ in casing.

---

## 3. SQL Medium Level: Views & Joins (Solutions Reference)

For reference, here are the expected queries for **SQL Medium Level: Views & Joins**:

1. **Task 1:** Create a VIEW named 'StudentEnrollments' that shows Student names and their Course titles.
   `SELECT Students.name, Courses.title FROM Students INNER JOIN Enrollments ON Students.id = Enrollments.student_id INNER JOIN Courses ON Enrollments.course_id = Courses.id`
2. **Task 2:** Select all columns from the 'StudentEnrollments' view.
   `SELECT * FROM StudentEnrollments`
3. **Task 3:** List all students and their enrollment date (if any) using a LEFT JOIN.
   `SELECT Students.name, Enrollments.enrollment_date FROM Students LEFT JOIN Enrollments ON Students.id = Enrollments.student_id`
4. **Task 4:** Find the total fee collected from all enrollments. (SUM of course fees)
   `SELECT SUM(Courses.fee) FROM Enrollments INNER JOIN Courses ON Enrollments.course_id = Courses.id`
5. **Task 5:** Find cities where more than 1 student resides. (Use HAVING)
   `SELECT city FROM Students GROUP BY city HAVING COUNT(*) > 1`
6. **Task 6:** Find students who joined after 'Ahmed Khan' (id=1).
   `SELECT * FROM Students WHERE joining_date > (SELECT joining_date FROM Students WHERE id = 1)`
7. **Task 7:** Show course titles and the number of students enrolled in each.
   `SELECT Courses.title, COUNT(Enrollments.student_id) FROM Courses LEFT JOIN Enrollments ON Courses.id = Enrollments.course_id GROUP BY Courses.title`
8. **Task 8:** Find the most expensive course title and its fee.
   `SELECT title, fee FROM Courses ORDER BY fee DESC LIMIT 1`
9. **Task 9:** Get the names of students enrolled in 'Python Basics'.
   `SELECT name FROM Students WHERE id IN (SELECT student_id FROM Enrollments WHERE course_id = 101)`
10. **Task 10:** Create a view named 'KarachiStudents' for students living in Karachi.
    `SELECT * FROM Students WHERE city = 'Karachi'`

---

## 4. SQL Basic Practical (Solutions Reference)

For reference, here are the expected queries for **SQL Basic Practical**:

1. `SELECT * FROM Students`
2. `SELECT * FROM Courses LIMIT 2`
3. `SELECT * FROM Students WHERE city = 'Karachi'`
4. `SELECT * FROM Students WHERE name LIKE 'A%'`
5. `SELECT * FROM Courses WHERE fee > 5000 AND duration_months < 12`
6. `SELECT * FROM Students WHERE city = 'Lahore' OR city = 'Islamabad'`
7. `SELECT * FROM Courses ORDER BY fee DESC`
8. `SELECT city, COUNT(*) FROM Students GROUP BY city`
9. `SELECT * FROM Students WHERE strftime('%m', joining_date) = '05'`
10. `SELECT Students.name, Courses.title FROM Students INNER JOIN Enrollments ON Students.id = Enrollments.student_id INNER JOIN Courses ON Enrollments.course_id = Courses.id`
