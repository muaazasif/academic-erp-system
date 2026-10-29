from app import app, db, Student, ExcelSkillsAssignment, ExcelSubmission
from werkzeug.security import generate_password_hash
from datetime import datetime

raw_data = """
504625	Marium Gul	Excel Skill 1: Formulas & Basics	6.77/10	67.70%	CLEAN	2026-04-16 15:50:32		2026-04-16 15:50:34
504622	Tabinda Fashar Samar 	Excel Skill 1: Formulas & Basics	7.43/10	74.30%	CLEAN	2026-04-16 16:50:13		2026-04-16 16:50:15
496875	Muhammad Nabeel	Excel Skill 1: Formulas & Basics	7.63/10	76.30%	CLEAN	2026-04-16 22:03:16		2026-04-16 22:03:17
497811	Muhammad Taha 	Excel Skill 1: Formulas & Basics	7.23/10	72.30%	CLEAN	2026-04-17 4:10:00		2026-04-17 4:10:01
59621	Hunzila Bin Naseem	Excel Skill 1: Formulas & Basics	7.63/10	76.30%	CLEAN	2026-05-13 10:55:21		2026-05-13 10:55:22
36575	Muhammad Zaman 	Excel Skill 1: Formulas & Basics	6.0/10	60.00%	CLEAN	2026-04-17 7:53:30		2026-04-17 7:53:31
499713	Ali Arman	Excel Skill 1: Formulas & Basics	2.8/10	28.00%	CLEAN	2026-06-18 18:48:23		2026-06-18 18:48:24
327312	Syed Raahim Ali	Excel Skill 1: Formulas & Basics	6.27/10	62.70%	CLEAN	2026-04-17 13:51:47		2026-04-17 13:51:48
496870	Mohammad Hammad	Excel Skill 1: Formulas & Basics	7.63/10	76.30%	CLEAN	2026-04-18 5:44:10		2026-04-18 5:44:11
488158	Ahmed khan 	Excel Skill 1: Formulas & Basics	7.3/10	73.00%	CLEAN	2026-04-18 9:05:37		2026-04-18 9:05:38
503639	Hafiz Rizwan Ahmed	Excel Skill 1: Formulas & Basics	3.32/10	33.20%	CLEAN	2026-04-18 9:11:11		2026-04-18 9:11:11
47541111	47541111	Excel Skill 1: Formulas & Basics	2.0/10	20.00%	CLEAN	2026-04-18 11:24:05		2026-04-18 11:24:06
475411	475411111	Excel Skill 1: Formulas & Basics	7.43/10	74.30%	CLEAN	2026-04-18 11:36:22		2026-04-18 11:36:22
504168	Muhammad Bazan Iftikhar	Excel Skill 1: Formulas & Basics	7.1/10	71.00%	CLEAN	2026-04-19 6:59:44		2026-04-19 6:59:45
503638	DANISH ALI	Excel Skill 1: Formulas & Basics	6.9/10	69.00%	CLEAN	2026-04-19 17:34:28		2026-04-19 17:34:29
508680	Kashif Akram	Excel Skill 1: Formulas & Basics	7.43/10	74.30%	CLEAN	2026-04-20 2:04:50		2026-04-20 2:04:52
327312	Syed Raahim Ali	Excel Skill 1: Formulas & Basics	6.27/10	62.70%	CLEAN	2026-04-20 6:20:09		2026-04-20 6:20:10
327312	Syed Raahim Ali	Excel Skill 1: Formulas & Basics	6.27/10	63%	CLEAN	2026-04-20 6:36:02		2026-04-20 6:36:02
504168	Muhammad Bazan Iftikhar	Excel Skill 2: Data Validation & Named Manager	5.5/10	55.00%	CLEAN	2026-04-20 7:10:39		2026-04-20 7:10:39
36575	Muhammad Zaman	Excel Skill 2: Data Validation & Named Manager	7.75/10	77.50%	CLEAN	2026-04-20 11:36:41		2026-04-20 11:36:41
497811	Muhammad Taha	Excel Skill 1: Formulas & Basics	7.23/10	72.30%	CLEAN	2026-04-20 12:25:45		2026-04-20 12:25:45
327312	Syed Raahim Ali	Excel Skill 2: Data Validation & Named Manager	8.25/10	82.50%	CLEAN	2026-04-20 14:54:28		2026-04-20 14:54:29
507312	Rao Abdul Razi 	Excel Skill 2: Data Validation & Named Manager	3.75/10	37.50%	CLEAN	2026-04-20 16:14:15		2026-04-20 16:14:15
508680	Kashif Akram	Excel Skill 1: Formulas & Basics	7.43/10	74.30%	CLEAN	2026-04-20 17:27:10		2026-04-20 17:27:11
475411	arshadalichandio70@gmail.com	Excel Skill 2: Data Validation & Named Manager	7.75/10	77.50%	CLEAN	2026-04-21 13:22:17		2026-04-21 13:22:18
475411	arshadalichandio70@gmail.com	Excel Skill 1: Formulas & Basics	7.43/10	74.30%	CLEAN	2026-04-21 13:22:43		2026-04-21 13:22:43
488158	Ahmed khan	Excel Skill 2: Data Validation & Named Manager	10.0/10	100.00%	CLEAN	2026-04-21 13:53:43		2026-04-21 13:53:45
504625	Marium Gul	Excel Skill 1: Formulas & Basics	7.63/10	76.30%	CLEAN	2026-04-21 14:52:46		2026-04-21 14:52:47
504622	Tabinda Fashar Samar	Excel Skill 1: Formulas & Basics	7.63/10	76.30%	CLEAN	2026-04-21 15:17:16		2026-04-21 15:17:17
499713	Ali Arman	Excel Skill 1: Formulas & Basics	6.6/10	66.00%	CLEAN	2026-04-21 18:40:12		2026-04-21 18:40:13
496952	Noor-jehan	Excel Skill 1: Formulas & Basics	7.1/10	71.00%	CLEAN	2026-04-22 3:30:32		2026-04-22 3:30:33
327312	Syed Raahim Ali	Excel Skill 1: Formulas & Basics	8.47/10	84.70%	CLEAN	2026-04-22 5:50:28		2026-04-22 5:50:29
327312	Syed Raahim Ali	Excel Skill 2: Data Validation & Named Manager	8.75/10	87.50%	CLEAN	2026-04-22 6:18:06		2026-04-22 6:18:07
507312	Rao Abdul Razi	Excel Skill 1: Formulas & Basics	9.63/10	96.30%	CLEAN	2026-04-22 13:26:24		2026-04-22 13:26:27
507312	Rao Abdul Razi	Excel Skill 2: Data Validation & Named Manager	5.95/10	59.50%	CLEAN	2026-04-22 13:47:39		2026-04-22 13:47:40
497811	Muhammad Taha	Excel Skill 2: Data Validation & Named Manager	10.0/10	100.00%	CLEAN	2026-04-22 16:47:44		2026-04-22 16:47:45
499713	Ali Arman	Excel Skill 2: Data Validation & Named Manager	5.25/10	52.50%	CLEAN	2026-06-18 18:48:40		2026-06-18 18:48:40
499713	Ali Arman	Excel Skill 1: Formulas & Basics	8.6/10	86.00%	CLEAN	2026-04-22 18:12:48		2026-04-22 18:12:48
493240	493240	Excel Skill 1: Formulas & Basics	4.85/10	48.50%	CLEAN	2026-04-23 18:40:35		2026-04-23 18:40:36
496875	Muhammad Nabeel	Excel Skill 2: Data Validation & Named Manager	7.75/10	77.50%	CLEAN	2026-04-23 21:05:28		2026-04-23 21:05:29
496875	Muhammad Nabeel	Excel Skill 1: Formulas & Basics	7.63/10	76.30%	CLEAN	2026-04-23 21:05:55		2026-04-23 21:05:55
503639	Hafiz Rizwan Ahmed	Excel Skill 1: Formulas & Basics	8.57/10	85.70%	CLEAN	2026-04-24 11:13:17		2026-04-24 11:13:18
387674	387674	Excel Skill 1: Formulas & Basics	9.63/10	96.30%	CLEAN	2026-04-24 16:06:12		2026-04-24 16:06:13
387674	387674	Excel Skill 2: Data Validation & Named Manager	9.0/10	90.00%	CLEAN	2026-04-24 16:50:04		2026-04-24 16:50:05
508680	Kashif Akram	Excel Skill 1: Formulas & Basics	9.43/10	94.30%	CLEAN	2026-04-24 19:00:15		2026-04-24 19:00:16
508680	Kashif Akram	Excel Skill 2: Data Validation & Named Manager	9.2/10	92.00%	CLEAN	2026-04-24 19:00:32		2026-04-24 19:00:33
503986	503986	Excel Skill 1: Formulas & Basics	5.98/10	59.80%	CLEAN	2026-04-25 5:49:55		2026-04-25 5:49:57
59621	Hunzila Bin Naseem	Excel Skill 2: Data Validation & Named Manager	7.75/10	77.50%	CLEAN	2026-05-13 10:55:48		2026-05-13 10:55:49
496870	Mohammad Hammad	Excel Skill 2: Data Validation & Named Manager	7.75/10	77.50%	CLEAN	2026-04-25 9:32:59		2026-04-25 9:33:00
496870	Mohammad Hammad	Excel Skill 1: Formulas & Basics	7.63/10	76.30%	CLEAN	2026-04-25 9:38:41		2026-04-25 9:38:41
504622	Tabinda Fashar Samar	Excel Skill 2: Data Validation & Named Manager	8.5/10	85.00%	CLEAN	2026-04-25 10:56:31		2026-04-25 10:56:31
503638	DANISH ALI	Excel Skill 2: Data Validation & Named Manager	10.0/10	100.00%	CLEAN	2026-04-25 12:28:47		2026-04-25 12:28:48
503639	Hafiz Rizwan Ahmed	Excel Skill 2: Data Validation & Named Manager	9.0/10	90.00%	CLEAN	2026-04-25 12:42:39		2026-04-25 12:42:39
504625	Marium Gul	Excel Skill 2: Data Validation & Named Manager	9.0/10	90.00%	CLEAN	2026-04-26 12:17:56		2026-04-26 12:17:57
504625	Marium Gul	Excel Skill 1: Formulas & Basics	9.63/10	96.30%	CLEAN	2026-04-26 12:18:19		2026-04-26 12:18:20
504622	Tabinda Fashar Samar	Excel Skill 1: Formulas & Basics	9.63/10	96.30%	CLEAN	2026-04-26 12:19:49		2026-04-26 12:19:49
341615	341615	Excel Skill 1: Formulas & Basics	9.1/10	91.00%	CLEAN	2026-04-26 16:29:29		2026-04-26 16:29:30
36575	Muhammad Zaman	Excel Skill 3: Data Cleaning & Power Query	6.88/10	68.80%	CLEAN	2026-04-27 12:13:35		2026-04-27 12:13:35
504622	Tabinda Fashar Samar	Excel Skill 2: Data Validation & Named Manager	9.0/10	90.00%	CLEAN	2026-04-27 15:43:15		2026-04-27 15:43:16
475411	arshadalichandio70@gmail.com	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	2026-04-28 15:51:21		2026-04-28 15:51:22
493240	493240	Excel Skill 1: Formulas & Basics	8.24/10	82.40%	CLEAN	2026-04-28 17:34:56		2026-04-28 17:34:57
497968	497968	Excel Skill 2: Data Validation & Named Manager	10.0/10	100.00%	CLEAN	2026-05-17 3:55:06		2026-05-17 3:55:07
497968	497968	Excel Skill 1: Formulas & Basics	9.43/10	94.30%	CLEAN	2026-05-17 3:54:46		2026-05-17 3:54:46
327312	Syed Raahim Ali	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	2026-04-29 4:13:25		2026-04-29 4:13:26
499249	499249	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	2026-05-07 21:14:35		2026-05-07 21:14:35
497811	Muhammad Taha	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	2026-04-29 7:20:07		2026-04-29 7:20:08
504168	Muhammad Bazan Iftikhar	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	2026-04-29 10:09:03		2026-04-29 10:09:04
50357555	50357555	Excel Skill 1: Formulas & Basics	7.81/10	78.10%	CLEAN	2026-04-29 16:12:25		2026-04-29 16:12:26
504625	Marium Gul	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	2026-04-29 16:48:59		2026-04-29 16:48:59
504622	Tabinda Fashar Samar	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	2026-04-29 16:52:03		2026-04-29 16:52:03
3431844	3431844	Excel Skill 1: Formulas & Basics	9.63/10	96.30%	CLEAN	2026-04-30 8:00:30		2026-04-30 8:00:31
485747	485747	Excel Skill 1: Formulas & Basics	9.43/10	94.30%	CLEAN	2026-05-15 21:25:37		2026-05-15 21:25:38
507312	Rao Abdul Razi	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	2026-04-30 11:48:12		2026-04-30 11:48:13
508680	Kashif Akram	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	2026-05-01 11:53:35		2026-05-01 11:53:37
499713	Ali Arman	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	2026-06-18 18:48:51		2026-06-18 18:48:52
499713	Ali Arman	Excel Skill 1: Formulas & Basics	8.6/10	86.00%	CLEAN	2026-05-01 16:08:19		2026-05-01 16:08:20
499713	Ali Arman	Excel Skill 2: Data Validation & Named Manager	5.25/10	52.50%	CLEAN	2026-05-01 16:08:29		2026-05-01 16:08:29
3431844	3431844	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	2026-05-01 18:28:39		2026-05-01 18:28:40
503808	503808	Excel Skill 1: Formulas & Basics	8.9/10	89.00%	CLEAN	2026-05-01 21:28:35		2026-05-01 21:28:36
503808	503808	Excel Skill 2: Data Validation & Named Manager	9.2/10	92.00%	CLEAN	2026-05-01 22:22:50		2026-05-01 22:22:51
503808	503808	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	2026-05-01 22:58:20		2026-05-01 22:58:21
485747	485747	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	2026-05-01 23:46:24		2026-05-01 23:46:25
503986	503986	Excel Skill 1: Formulas & Basics	5.98/10	59.80%	CLEAN	2026-05-02 4:43:47		2026-05-02 4:43:47
503986	503986	Excel Skill 2: Data Validation & Named Manager	10.0/10	100.00%	CLEAN	2026-05-02 5:16:23		2026-05-02 5:16:24
496952	Noor-jehan	Excel Skill 1: Formulas & Basics	9.1/10	91.00%	CLEAN	2026-05-02 8:45:04		2026-05-02 8:45:04
497968	497968	Excel Skill 1: Formulas & Basics	9.43/10	94.30%	CLEAN	2026-05-03 18:38:41		2026-05-03 18:38:42
497968	497968	Excel Skill 2: Data Validation & Named Manager	10.0/10	100.00%	CLEAN	2026-05-03 18:39:11		2026-05-03 18:39:12
497968	497968	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	2026-05-17 3:55:24		2026-05-17 3:55:24
503575	afia khurram	Excel Skill 2: Data Validation & Named Manager	9.2/10	92.00%	CLEAN	2026-05-06 15:08:54		2026-05-06 15:08:55
503575	afia khurram	Excel Skill 1: Formulas & Basics	7.81/10	78.10%	CLEAN	2026-05-06 15:09:50		2026-05-06 15:09:50
503575	afia khurram	Excel Skill 3: Data Cleaning & Power Query	9.98/10	99.80%	CLEAN	2026-05-06 15:06:58		2026-05-06 15:07:00
493240	493240	Excel Skill 2: Data Validation & Named Manager	0/10	0.00%	CLEAN	2026-05-07 18:04:25		2026-05-07 18:04:26
493240	493240	Excel Skill 3: Data Cleaning & Power Query	0/10	0.00%	CLEAN	2026-05-07 18:08:23		2026-05-07 18:08:24
499249	499249	Excel Skill 1: Formulas & Basics	8.3/10	83.00%	CLEAN	2026-05-07 21:14:19		2026-05-07 21:14:20
499249	499249	Excel Skill 2: Data Validation & Named Manager	5.5/10	55.00%	CLEAN	2026-05-07 22:01:05		2026-05-07 22:01:06
4932400	4932400	Excel Skill 2: Data Validation & Named Manager	8.1/10	81.00%	CLEAN	2026-05-08 6:17:18		2026-05-08 6:17:18
4932400	4932400	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	2026-05-08 6:18:34		2026-05-08 6:18:34
387452	387452	Excel Skill 1: Formulas & Basics	9.3/10	93.00%	CLEAN	2026-05-09 12:48:49		2026-05-09 12:48:50
387452	387452	Excel Skill 2: Data Validation & Named Manager	8.2/10	82.00%	CLEAN	2026-06-01 19:47:13		2026-06-01 19:47:13
387452	387452	Excel Skill 3: Data Cleaning & Power Query	6.88/10	68.80%	CLEAN	2026-06-01 19:46:55		2026-06-01 19:46:56
485747	485747	Excel Skill 2: Data Validation & Named Manager	8.1/10	81.00%	CLEAN	2026-05-15 21:25:58		2026-05-15 21:25:59
496952	Noor-jehan	Excel Skill 2: Data Validation & Named Manager	9/10	90%	CLEAN	2026-05-10 10:28:28		2026-05-10 10:28:29
123	123	Excel Skill 2: Data Validation & Named Manager	10.0/10	100.00%	CLEAN	2026-09-28 0:42:42		ProductList: Correct
491981	491981	Excel Skill 1: Formulas & Basics	9.13/10	91.30%	CLEAN	2026-05-12 17:34:41		2026-05-12 17:34:43
491981	491981	Excel Skill 3: Data Cleaning & Power Query	2.5/10	25.00%	CLEAN	2026-05-12 22:50:11		2026-05-12 22:50:12
496870	Mohammad Hammad	Excel Skill 3: Data Cleaning & Power Query	0/10	0.00%	CLEAN	2026-05-13 11:06:16		2026-05-13 11:06:17
491981	491981	Excel Skill 1: Formulas & Basics	9.13/10	91.30%	CLEAN	2026-05-19 16:50:58		2026-05-19 16:50:59
491981	491981	Excel Skill 3: Data Cleaning & Power Query	4.48/10	44.80%	CLEAN	2026-05-13 16:47:59		2026-05-13 16:47:59
491981	491981	Excel Skill 2: Data Validation & Named Manager	6.25/10	62.50%	CLEAN	2026-05-14 17:43:28		2026-05-14 17:43:29
491981	491981	Excel Skill 2: Data Validation & Named Manager	6.25/10	62.50%	CLEAN	2026-05-19 16:51:12		2026-05-19 16:51:12
496952	Noor-jehan	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	2026-05-15 3:44:23		2026-05-15 3:44:24
488158	Ahmed khan	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	2026-05-20 16:28:31		2026-05-20 16:28:33
123	123	Excel Skill 4: Advanced LOOKUP & Aggregation	8.45/10	84.50%	CLEAN	2026-06-19 5:59:59		2026-06-19 5:59:59
387452	387452	Excel Skill 4: Advanced LOOKUP & Aggregation	5.0/10	50.00%	CLEAN	2026-06-01 19:45:31		2026-06-01 19:45:33
2007	Aazmir	Excel Skill 1: Formulas & Basics	2.0/10	20.00%	CLEAN	2026-06-02 6:44:42		2026-06-02 6:44:43
123	123	Excel Skill 1: Formulas & Basics	0.5/10	5.00%	CLEAN	2026-10-10 1:00:36	Q1: Incorrect
2007	Aazmir	Excel Skill 1: Formulas & Basics	9.3/10	93.00%	CLEAN	2026-06-02 6:55:45		2026-06-02 6:55:45
2007	Aazmir	Excel Skill 2: Data Validation & Named Manager	10.0/10	100.00%	CLEAN	2026-06-02 9:15:39		2026-06-02 9:15:40
504168	Muhammad Bazan Iftikhar	Excel Skill 4: Advanced LOOKUP & Aggregation	9.25/10	92.50%	CLEAN	2026-06-02 11:13:25		2026-06-02 11:13:27
2007	Aazmir	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	2026-06-03 10:56:06		2026-06-03 10:56:07
444	Afsheen	Excel Skill 1: Formulas & Basics	3.5/10	35.00%	CLEAN	2026-06-05 3:36:12		2026-06-05 3:36:13
4444	4444	Excel Skill 1: Formulas & Basics	3.5/10	35.00%	CLEAN	2026-09-29 18:07:10	Q1: Correct
444	Afsheen	Excel Skill 2: Data Validation & Named Manager	10.0/10	100.00%	CLEAN	2026-06-05 10:37:45		2026-06-05 10:37:46
4444	Prof Afsheen	Excel Skill 2: Data Validation & Named Manager	10.0/10	100.00%	CLEAN	2026-06-05 10:38:51		2026-06-05 10:38:52
444	Afsheen	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	2026-06-08 10:34:27		2026-06-08 10:34:28
4444	Prof Afsheen	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	2026-06-08 10:43:14		2026-06-08 10:43:14
485747	485747	Excel Skill 4: Advanced LOOKUP & Aggregation	4.2/10	42.00%	CLEAN	2026-06-08 19:22:52		2026-06-08 19:22:54
444	Afsheen	Excel Skill 4: Advanced LOOKUP & Aggregation	9.0/10	90.00%	CLEAN	2026-06-09 7:40:29		2026-06-09 7:40:30
4444	Prof Afsheen	Excel Skill 4: Advanced LOOKUP & Aggregation	9.0/10	90.00%	CLEAN	2026-06-09 7:42:37		2026-06-09 7:42:38
503575	afia khurram	Excel Skill 4: Advanced LOOKUP & Aggregation	8.0/10	80.00%	CLEAN	2026-06-11 16:01:43		2026-06-11 16:01:45
499713	Ali Arman	Excel Skill 4: Advanced LOOKUP & Aggregation	9.0/10	90.00%	CLEAN	2026-06-18 19:20:20		2026-06-18 19:20:21
504625	Marium Gul	Excel Skill 4: Advanced LOOKUP & Aggregation	10.0/10	100.00%	CLEAN	2026-06-19 5:45:46		2026-06-19 5:45:46
504622	Tabinda Fashar Samar	Excel Skill 4: Advanced LOOKUP & Aggregation	10.0/10	100.00%	CLEAN	2026-06-19 5:46:33		2026-06-19 5:46:33
497811	Muhammad Taha	Excel Skill 4: Advanced LOOKUP & Aggregation	9.0/10	90.00%	CLEAN	2026-06-24 11:52:21		2026-06-24 11:52:22
496952	Noor-jehan	Excel Skill 4: Advanced LOOKUP & Aggregation	9.2/10	92.00%	CLEAN	2026-07-04 4:50:07		2026-07-04 4:50:08
327312	Syed Raahim Ali	Excel Skill 4: Advanced LOOKUP & Aggregation	10.0/10	100.00%	CLEAN	2026-07-04 10:12:50		2026-07-04 10:12:51
383068	383068	Excel Skill 1: Formulas & Basics	7.3/10	73.00%	CLEAN	2026-07-07 16:20:05		2026-07-07 16:20:07
383068	383068	Excel Skill 2: Data Validation & Named Manager	4.1/10	41.00%	CLEAN	2026-07-08 12:47:58		2026-07-08 12:48:00
383068	383068	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	2026-07-09 13:36:09		2026-07-09 13:36:11
383068	383068	Excel Skill 4: Advanced LOOKUP & Aggregation	7.5/10	75.00%	CLEAN	2026-07-10 13:01:46		2026-07-10 13:01:47
475411	arshadalichandio70@gmail.com	Excel Skill 4: Advanced LOOKUP & Aggregation	5.8/10	58.00%	CLEAN	2026-07-20 14:11:42		2026-07-20 14:11:42
507312	Rao Abdul Razi	Excel Skill 4: Advanced LOOKUP & Aggregation	10.0/10	100.00%	CLEAN	2026-07-22 14:07:16		2026-07-22 14:07:17
505493	505493	Excel Skill 1: Formulas & Basics	0/10	0.00%	CLEAN	2026-08-03 17:21:03		2026-08-03 17:21:04
505493	505493	Excel Skill 2: Data Validation & Named Manager	0/10	0%	CLEAN	2026-08-04 8:09:14		2026-08-04 8:09:16
505493	505493	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	2026-08-04 10:22:21		2026-08-04 10:22:22
5054933	5054933	Excel Skill 1: Formulas & Basics	8.9/10	89.00%	CLEAN	2026-08-04 11:04:11		2026-08-04 11:04:11
5054933	5054933	Excel Skill 2: Data Validation & Named Manager	0/10	0%	CLEAN	2026-08-04 11:05:15		2026-08-04 11:05:15
564453	Duksh Kumar	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	2026-10-01 10:15:28	Q1: Correct
555402	Alisha Amjad	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	2026-10-03 9:34:17	Q1: Correct
545050	Hammad Sharif	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	2.5/10	25.00%	CLEAN	2026-10-17 18:06:54	Q1: Incorrect
"""

with app.app_context():
    # Cache assignments
    assignments = {a.title.strip(): a.id for a in ExcelSkillsAssignment.query.all()}
    print(f"Loaded {len(assignments)} Excel assignments.")

    lines = [l.strip() for l in raw_data.strip().split('\n') if l.strip()]
    print(f"Parsed {len(lines)} submission lines.")

    count = 0
    for line in lines:
        parts = line.split('\t')
        if len(parts) < 7:
            continue
        sid = parts[0].strip()
        name = parts[1].strip()
        assign_title = parts[2].strip()
        score_str = parts[3].strip() # e.g. "6.77/10" or "5.0/5"
        pct_str = parts[4].strip() # e.g. "67.70%"
        status = parts[5].strip()
        submitted_at_str = parts[6].strip()
        grade_details = parts[7].strip() if len(parts) > 7 else ""

        # Ensure student exists
        student = Student.query.filter_by(student_id=sid).first()
        if not student:
            student = Student(
                student_id=sid,
                name=name or f"Student {sid}",
                batch="Bano Qabil 5.0 Fall",
                gender="Female",
                password_hash=generate_password_hash("student123")
            )
            db.session.add(student)
            db.session.flush()

        # Find assignment ID
        assign_id = assignments.get(assign_title)
        if not assign_id:
            # Try fuzzy match
            for title, aid in assignments.items():
                if assign_title.lower() in title.lower() or title.lower() in assign_title.lower():
                    assign_id = aid
                    break
        if not assign_id:
            print(f"Warning: Assignment '{assign_title}' not found.")
            continue

        # Parse score
        try:
            score = float(score_str.split('/')[0])
        except Exception:
            score = 0.0

        # Parse percentage
        try:
            pct = float(pct_str.replace('%', '').strip())
        except Exception:
            pct = 0.0

        # Parse datetime
        try:
            sub_at = datetime.strptime(submitted_at_str, "%Y-%m-%d %H:%M:%S")
        except Exception:
            sub_at = datetime.utcnow()

        sub = ExcelSubmission(
            assignment_id=assign_id,
            student_id=sid,
            score=score,
            percentage=pct,
            grade_details=grade_details,
            status='graded',
            submitted_at=sub_at
        )
        db.session.add(sub)
        count += 1
        if count % 50 == 0:
            db.session.commit()

    db.session.commit()
    print(f"✅ Successfully restored {count} Excel submission records!")
    print(f"Total excel submissions in DB: {ExcelSubmission.query.count()}")
