import csv
import io
from app import app, db, Student, ExcelSkillsAssignment, ExcelSubmission
from werkzeug.security import generate_password_hash
from datetime import datetime

raw_data = """
504625	Marium Gul	Excel Skill 1: Formulas & Basics	6.77/10	67.70%	CLEAN	4/16/2026 15:50	
504622	Tabinda Fashar Samar	Excel Skill 1: Formulas & Basics	7.43/10	74.30%	CLEAN	4/16/2026 16:50	
496875	Muhammad Nabeel	Excel Skill 1: Formulas & Basics	7.63/10	76.30%	CLEAN	4/16/2026 22:03	
497811	Muhammad Taha	Excel Skill 1: Formulas & Basics	7.23/10	72.30%	CLEAN	4/17/2026 4:10	
59621	Hunzila Bin Naseem	Excel Skill 1: Formulas & Basics	7.63/10	76.30%	CLEAN	5/13/2026 10:55	
36575	Muhammad Zaman	Excel Skill 1: Formulas & Basics	6.0/10	60.00%	CLEAN	4/17/2026 7:53	
499713	Ali Arman	Excel Skill 1: Formulas & Basics	2.8/10	28.00%	CLEAN	6/18/2026 18:48	
327312	Syed Raahim Ali	Excel Skill 1: Formulas & Basics	6.27/10	62.70%	CLEAN	4/17/2026 13:51	
496870	Mohammad Hammad	Excel Skill 1: Formulas & Basics	7.63/10	76.30%	CLEAN	4/18/2026 5:44	
488158	Ahmed khan	Excel Skill 1: Formulas & Basics	7.3/10	73.00%	CLEAN	4/18/2026 9:05	
503639	Hafiz Rizwan Ahmed	Excel Skill 1: Formulas & Basics	3.32/10	33.20%	CLEAN	4/18/2026 9:11	
47541111	47541111	Excel Skill 1: Formulas & Basics	2.0/10	20.00%	CLEAN	4/18/2026 11:24	
475411	475411111	Excel Skill 1: Formulas & Basics	7.43/10	74.30%	CLEAN	4/18/2026 11:36	
504168	Muhammad Bazan Iftikhar	Excel Skill 1: Formulas & Basics	7.1/10	71.00%	CLEAN	4/19/2026 6:59	
503638	DANISH ALI	Excel Skill 1: Formulas & Basics	6.9/10	69.00%	CLEAN	4/19/2026 17:34	
508680	Kashif Akram	Excel Skill 1: Formulas & Basics	7.43/10	74.30%	CLEAN	4/20/2026 2:04	
327312	Syed Raahim Ali	Excel Skill 1: Formulas & Basics	6.27/10	62.70%	CLEAN	4/20/2026 6:20	
327312	Syed Raahim Ali	Excel Skill 1: Formulas & Basics	6.27/10	63%	CLEAN	4/20/2026 6:36	
504168	Muhammad Bazan Iftikhar	Excel Skill 2: Data Validation & Named Manager	5.5/10	55.00%	CLEAN	4/20/2026 7:10	
36575	Muhammad Zaman	Excel Skill 2: Data Validation & Named Manager	7.75/10	77.50%	CLEAN	4/20/2026 11:36	
497811	Muhammad Taha	Excel Skill 1: Formulas & Basics	7.23/10	72.30%	CLEAN	4/20/2026 12:25	
327312	Syed Raahim Ali	Excel Skill 2: Data Validation & Named Manager	8.25/10	82.50%	CLEAN	4/20/2026 14:54	
507312	Rao Abdul Razi	Excel Skill 2: Data Validation & Named Manager	3.75/10	37.50%	CLEAN	4/20/2026 16:14	
508680	Kashif Akram	Excel Skill 1: Formulas & Basics	7.43/10	74.30%	CLEAN	4/20/2026 17:27	
475411	arshadalichandio70 @gmail.com	Excel Skill 2: Data Validation & Named Manager	7.75/10	77.50%	CLEAN	4/21/2026 13:22	
475411	arshadalichandio70 @gmail.com	Excel Skill 1: Formulas & Basics	7.43/10	74.30%	CLEAN	4/21/2026 13:22	
488158	Ahmed khan	Excel Skill 2: Data Validation & Named Manager	10.0/10	100.00%	CLEAN	4/21/2026 13:53	
504625	Marium Gul	Excel Skill 1: Formulas & Basics	7.63/10	76.30%	CLEAN	4/21/2026 14:52	
504622	Tabinda Fashar Samar	Excel Skill 1: Formulas & Basics	7.63/10	76.30%	CLEAN	4/21/2026 15:17	
499713	Ali Arman	Excel Skill 1: Formulas & Basics	6.6/10	66.00%	CLEAN	4/21/2026 18:40	
496952	Noor-jehan	Excel Skill 1: Formulas & Basics	7.1/10	71.00%	CLEAN	4/22/2026 3:30	
327312	Syed Raahim Ali	Excel Skill 1: Formulas & Basics	8.47/10	84.70%	CLEAN	4/22/2026 5:50	
327312	Syed Raahim Ali	Excel Skill 2: Data Validation & Named Manager	8.75/10	87.50%	CLEAN	4/22/2026 6:18	
507312	Rao Abdul Razi	Excel Skill 1: Formulas & Basics	9.63/10	96.30%	CLEAN	4/22/2026 13:26	
507312	Rao Abdul Razi	Excel Skill 2: Data Validation & Named Manager	5.95/10	59.50%	CLEAN	4/22/2026 13:47	
497811	Muhammad Taha	Excel Skill 2: Data Validation & Named Manager	10.0/10	100.00%	CLEAN	4/22/2026 16:47	
499713	Ali Arman	Excel Skill 2: Data Validation & Named Manager	5.25/10	52.50%	CLEAN	6/18/2026 18:48	
499713	Ali Arman	Excel Skill 1: Formulas & Basics	8.6/10	86.00%	CLEAN	4/22/2026 18:12	
493240	493240	Excel Skill 1: Formulas & Basics	4.85/10	48.50%	CLEAN	4/23/2026 18:40	
496875	Muhammad Nabeel	Excel Skill 2: Data Validation & Named Manager	7.75/10	77.50%	CLEAN	4/23/2026 21:05	
496875	Muhammad Nabeel	Excel Skill 1: Formulas & Basics	7.63/10	76.30%	CLEAN	4/23/2026 21:05	
503639	Hafiz Rizwan Ahmed	Excel Skill 1: Formulas & Basics	8.57/10	85.70%	CLEAN	4/24/2026 11:13	
387674	387674	Excel Skill 1: Formulas & Basics	9.63/10	96.30%	CLEAN	4/24/2026 16:06	
387674	387674	Excel Skill 2: Data Validation & Named Manager	9.0/10	90.00%	CLEAN	4/24/2026 16:50	
508680	Kashif Akram	Excel Skill 1: Formulas & Basics	9.43/10	94.30%	CLEAN	4/24/2026 19:00	
508680	Kashif Akram	Excel Skill 2: Data Validation & Named Manager	9.2/10	92.00%	CLEAN	4/24/2026 19:00	
503986	503986	Excel Skill 1: Formulas & Basics	5.98/10	59.80%	CLEAN	4/25/2026 5:49	
59621	Hunzila Bin Naseem	Excel Skill 2: Data Validation & Named Manager	7.75/10	77.50%	CLEAN	5/13/2026 10:55	
496870	Mohammad Hammad	Excel Skill 2: Data Validation & Named Manager	7.75/10	77.50%	CLEAN	4/25/2026 9:32	
496870	Mohammad Hammad	Excel Skill 1: Formulas & Basics	7.63/10	76.30%	CLEAN	4/25/2026 9:38	
504622	Tabinda Fashar Samar	Excel Skill 2: Data Validation & Named Manager	8.5/10	85.00%	CLEAN	4/25/2026 10:56	
503638	DANISH ALI	Excel Skill 2: Data Validation & Named Manager	10.0/10	100.00%	CLEAN	4/25/2026 12:28	
503639	Hafiz Rizwan Ahmed	Excel Skill 2: Data Validation & Named Manager	9.0/10	90.00%	CLEAN	4/25/2026 12:42	
504625	Marium Gul	Excel Skill 2: Data Validation & Named Manager	9.0/10	90.00%	CLEAN	4/26/2026 12:17	
504625	Marium Gul	Excel Skill 1: Formulas & Basics	9.63/10	96.30%	CLEAN	4/26/2026 12:18	
504622	Tabinda Fashar Samar	Excel Skill 1: Formulas & Basics	9.63/10	96.30%	CLEAN	4/26/2026 12:19	
341615	341615	Excel Skill 1: Formulas & Basics	9.1/10	91.00%	CLEAN	4/26/2026 16:29	
36575	Muhammad Zaman	Excel Skill 3: Data Cleaning & Power Query	6.88/10	68.80%	CLEAN	4/27/2026 12:13	
504622	Tabinda Fashar Samar	Excel Skill 2: Data Validation & Named Manager	9.0/10	90.00%	CLEAN	4/27/2026 15:43	
475411	arshadalichandio70 @gmail.com	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	4/28/2026 15:51	
493240	493240	Excel Skill 1: Formulas & Basics	8.24/10	82.40%	CLEAN	4/28/2026 17:34	
497968	497968	Excel Skill 2: Data Validation & Named Manager	10.0/10	100.00%	CLEAN	5/17/2026 3:55	
497968	497968	Excel Skill 1: Formulas & Basics	9.43/10	94.30%	CLEAN	5/17/2026 3:54	
327312	Syed Raahim Ali	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	4/29/2026 4:13	
499249	499249	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	5/7/2026 21:14	
497811	Muhammad Taha	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	4/29/2026 7:20	
504168	Muhammad Bazan Iftikhar	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	4/29/2026 10:09	
50357555	50357555	Excel Skill 1: Formulas & Basics	7.81/10	78.10%	CLEAN	4/29/2026 16:12	
504625	Marium Gul	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	4/29/2026 16:48	
504622	Tabinda Fashar Samar	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	4/29/2026 16:52	
3431844	3431844	Excel Skill 1: Formulas & Basics	9.63/10	96.30%	CLEAN	4/30/2026 8:00	
485747	485747	Excel Skill 1: Formulas & Basics	9.43/10	94.30%	CLEAN	5/15/2026 21:25	
507312	Rao Abdul Razi	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	4/30/2026 11:48	
508680	Kashif Akram	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	5/1/2026 11:53	
499713	Ali Arman	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	6/18/2026 18:48	
499713	Ali Arman	Excel Skill 1: Formulas & Basics	8.6/10	86.00%	CLEAN	5/1/2026 16:08	
499713	Ali Arman	Excel Skill 2: Data Validation & Named Manager	5.25/10	52.50%	CLEAN	5/1/2026 16:08	
3431844	3431844	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	5/1/2026 18:28	
503808	503808	Excel Skill 1: Formulas & Basics	8.9/10	89.00%	CLEAN	5/1/2026 21:28	
503808	503808	Excel Skill 2: Data Validation & Named Manager	9.2/10	92.00%	CLEAN	5/1/2026 22:22	
503808	503808	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	5/1/2026 22:58	
485747	485747	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	5/1/2026 23:46	
503986	503986	Excel Skill 1: Formulas & Basics	5.98/10	59.80%	CLEAN	5/2/2026 4:43	
503986	503986	Excel Skill 2: Data Validation & Named Manager	10.0/10	100.00%	CLEAN	5/2/2026 5:16	
496952	Noor-jehan	Excel Skill 1: Formulas & Basics	9.1/10	91.00%	CLEAN	5/2/2026 8:45	
497968	497968	Excel Skill 1: Formulas & Basics	9.43/10	94.30%	CLEAN	5/3/2026 18:38	
497968	497968	Excel Skill 2: Data Validation & Named Manager	10.0/10	100.00%	CLEAN	5/3/2026 18:39	
497968	497968	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	5/17/2026 3:55	
503575	afia khurram	Excel Skill 2: Data Validation & Named Manager	9.2/10	92.00%	CLEAN	5/6/2026 15:08	
503575	afia khurram	Excel Skill 1: Formulas & Basics	7.81/10	78.10%	CLEAN	5/6/2026 15:09	
503575	afia khurram	Excel Skill 3: Data Cleaning & Power Query	9.98/10	99.80%	CLEAN	5/6/2026 15:06	
493240	493240	Excel Skill 2: Data Validation & Named Manager	0/10	0.00%	CLEAN	5/7/2026 18:04	
493240	493240	Excel Skill 3: Data Cleaning & Power Query	0/10	0.00%	CLEAN	5/7/2026 18:08	
499249	499249	Excel Skill 1: Formulas & Basics	8.3/10	83.00%	CLEAN	5/7/2026 21:14	
499249	499249	Excel Skill 2: Data Validation & Named Manager	5.5/10	55.00%	CLEAN	5/7/2026 22:01	
4932400	4932400	Excel Skill 2: Data Validation & Named Manager	8.1/10	81.00%	CLEAN	5/8/2026 6:17	
4932400	4932400	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	5/8/2026 6:18	
387452	387452	Excel Skill 1: Formulas & Basics	9.3/10	93.00%	CLEAN	5/9/2026 12:48	
387452	387452	Excel Skill 2: Data Validation & Named Manager	8.2/10	82.00%	CLEAN	6/1/2026 19:47	
387452	387452	Excel Skill 3: Data Cleaning & Power Query	6.88/10	68.80%	CLEAN	6/1/2026 19:46	
485747	485747	Excel Skill 2: Data Validation & Named Manager	8.1/10	81.00%	CLEAN	5/15/2026 21:25	
496952	Noor-jehan	Excel Skill 2: Data Validation & Named Manager	9/10	90%	CLEAN	5/10/2026 10:28	
123	123	Excel Skill 2: Data Validation & Named Manager	10.0/10	100.00%	CLEAN	9/28/2026 0:42	"ProductList: Correct
PriceList: Correct
CategoryList: Correct
C5: Correct
C7: Correct
Named Range - Electronics: Correct
Named Range - Furniture: Correct
C10: Correct
C11: Correct
C5 (Whole Number): Correct
C7 (Date): Correct
C9 (Text Length): Correct"
491981	491981	Excel Skill 1: Formulas & Basics	9.13/10	91.30%	CLEAN	5/12/2026 17:34	
491981	491981	Excel Skill 3: Data Cleaning & Power Query	2.5/10	25.00%	CLEAN	5/12/2026 22:50	
496870	Mohammad Hammad	Excel Skill 3: Data Cleaning & Power Query	0/10	0.00%	CLEAN	5/13/2026 11:06	
491981	491981	Excel Skill 1: Formulas & Basics	9.13/10	91.30%	CLEAN	5/19/2026 16:50	
491981	491981	Excel Skill 3: Data Cleaning & Power Query	4.48/10	44.80%	CLEAN	5/13/2026 16:47	
491981	491981	Excel Skill 2: Data Validation & Named Manager	6.25/10	62.50%	CLEAN	5/14/2026 17:43	
491981	491981	Excel Skill 2: Data Validation & Named Manager	6.25/10	62.50%	CLEAN	5/19/2026 16:51	
496952	Noor-jehan	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	5/15/2026 3:44	
488158	Ahmed khan	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	5/20/2026 16:28	
123	123	Excel Skill 4: Advanced LOOKUP & Aggregation	8.45/10	84.50%	CLEAN	6/19/2026 5:59	
387452	387452	Excel Skill 4: Advanced LOOKUP & Aggregation	5.0/10	50.00%	CLEAN	6/1/2026 19:45	
2007	Aazmir	Excel Skill 1: Formulas & Basics	2.0/10	20.00%	CLEAN	6/2/2026 6:44	
123	123	Excel Skill 1: Formulas & Basics	0.5/10	5.00%	CLEAN	10/10/2026 1:00	"Q1: Incorrect (No formula found.)
Q2: Incorrect (No formula found.)
Q3: Incorrect (No formula found.)
Q4: Incorrect (No formula found.)
Q5: Incorrect (No formula found.)
Q6: Incorrect (No formula found.)
Q7: Incorrect (No formula found.)
Q8: Incorrect (No formula found.)
Q9: Incorrect (No formula found.)
Q10: Incorrect (No formula found.)
Q11: Incorrect (No formula found.)
Q12: Incorrect (No formula found.)
Q13: Incorrect (No formula found.)
Q14: Incorrect (No formula found.)
Q15: Incorrect (No formula found.)
Q16: Correct
Q17: Incorrect (No formula found.)
Q18: Incorrect (No formula found.)
Q19: Incorrect (No formula found.)
Q20: Incorrect (No formula found.)"
2007	Aazmir	Excel Skill 1: Formulas & Basics	9.3/10	93.00%	CLEAN	6/2/2026 6:55	
2007	Aazmir	Excel Skill 2: Data Validation & Named Manager	10.0/10	100.00%	CLEAN	6/2/2026 9:15	
504168	Muhammad Bazan Iftikhar	Excel Skill 4: Advanced LOOKUP & Aggregation	9.25/10	92.50%	CLEAN	6/2/2026 11:13	
2007	Aazmir	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	6/3/2026 10:56	
444	Afsheen	Excel Skill 1: Formulas & Basics	3.5/10	35.00%	CLEAN	6/5/2026 3:36	
4444	4444	Excel Skill 1: Formulas & Basics	3.5/10	35.00%	CLEAN	9/29/2026 18:07	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Incorrect (table_range is incorrect. | return_column is incorrect.)
Q5: Correct
Q6: Incorrect (No formula found.)
Q7: Incorrect (No formula found.)
Q8: Incorrect (No formula found.)
Q9: Incorrect (No formula found.)
Q10: Incorrect (No formula found.)
Q11: Incorrect (No formula found.)
Q12: Incorrect (No formula found.)
Q13: Incorrect (No formula found.)
Q14: Incorrect (No formula found.)
Q15: Incorrect (No formula found.)
Q16: Incorrect (No formula found.)
Q17: Correct
Q18: Incorrect (No formula found.)
Q19: Correct
Q20: Correct"
444	Afsheen	Excel Skill 2: Data Validation & Named Manager	10.0/10	100.00%	CLEAN	6/5/2026 10:37	
4444	Prof Afsheen	Excel Skill 2: Data Validation & Named Manager	10.0/10	100.00%	CLEAN	6/5/2026 10:38	
444	Afsheen	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	6/8/2026 10:34	
4444	Prof Afsheen	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	6/8/2026 10:43	
485747	485747	Excel Skill 4: Advanced LOOKUP & Aggregation	4.2/10	42.00%	CLEAN	6/8/2026 19:22	
444	Afsheen	Excel Skill 4: Advanced LOOKUP & Aggregation	9.0/10	90.00%	CLEAN	6/9/2026 7:40	
4444	Prof Afsheen	Excel Skill 4: Advanced LOOKUP & Aggregation	9.0/10	90.00%	CLEAN	6/9/2026 7:42	
503575	afia khurram	Excel Skill 4: Advanced LOOKUP & Aggregation	8.0/10	80.00%	CLEAN	6/11/2026 16:01	
499713	Ali Arman	Excel Skill 4: Advanced LOOKUP & Aggregation	9.0/10	90.00%	CLEAN	6/18/2026 19:20	
504625	Marium Gul	Excel Skill 4: Advanced LOOKUP & Aggregation	10.0/10	100.00%	CLEAN	6/19/2026 5:45	
504622	Tabinda Fashar Samar	Excel Skill 4: Advanced LOOKUP & Aggregation	10.0/10	100.00%	CLEAN	6/19/2026 5:46	
497811	Muhammad Taha	Excel Skill 4: Advanced LOOKUP & Aggregation	9.0/10	90.00%	CLEAN	6/24/2026 11:52	
496952	Noor-jehan	Excel Skill 4: Advanced LOOKUP & Aggregation	9.2/10	92.00%	CLEAN	7/4/2026 4:50	
327312	Syed Raahim Ali	Excel Skill 4: Advanced LOOKUP & Aggregation	10.0/10	100.00%	CLEAN	7/4/2026 10:12	
383068	383068	Excel Skill 1: Formulas & Basics	7.3/10	73.00%	CLEAN	7/7/2026 16:20	
383068	383068	Excel Skill 2: Data Validation & Named Manager	4.1/10	41.00%	CLEAN	7/8/2026 12:47	
383068	383068	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	7/9/2026 13:36	
383068	383068	Excel Skill 4: Advanced LOOKUP & Aggregation	7.5/10	75.00%	CLEAN	7/10/2026 13:01	
475411	arshadalichandio70 @gmail.com	Excel Skill 4: Advanced LOOKUP & Aggregation	5.8/10	58.00%	CLEAN	7/20/2026 14:11	
507312	Rao Abdul Razi	Excel Skill 4: Advanced LOOKUP & Aggregation	10.0/10	100.00%	CLEAN	7/22/2026 14:07	
505493	505493	Excel Skill 1: Formulas & Basics	0/10	0.00%	CLEAN	8/3/2026 17:21	
505493	505493	Excel Skill 2: Data Validation & Named Manager	0/10	0%	CLEAN	8/4/2026 8:09	
505493	505493	Excel Skill 3: Data Cleaning & Power Query	7.48/10	74.80%	CLEAN	8/4/2026 10:22	
5054933	5054933	Excel Skill 1: Formulas & Basics	8.9/10	89.00%	CLEAN	8/4/2026 11:04	
5054933	5054933	Excel Skill 2: Data Validation & Named Manager	0/10	0%	CLEAN	8/4/2026 11:05	
564453	Duksh Kumar	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/1/2026 10:15	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
555402	Alisha Amjad	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/3/2026 9:34	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
545050	Hammad Sharif	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	2.5/10	25.00%	CLEAN	10/17/2026 18:06	"Q1: Incorrect (Wrong lookup value: A53. Expected E050.)
Q2: Correct
Q3: Correct
Q4: Incorrect (IF ke exactly 3 arguments hone chahiye.)"
582660	Shaheer Asif	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/3/2026 18:28	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
541501	Esha Ansari	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/2/2026 3:26	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
567003	Kamran Khalid	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/5/2026 5:37	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
483031	Roma Arshad	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/2/2026 9:29	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
571660	Zainab	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	3.75/10	37.50%	CLEAN	10/12/2026 10:51	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Incorrect (Wrong condition: E4>45000. Expected C108>45000.)"
554426	Muhammad Zubair	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	9/28/2026 13:31	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
559189	Hafiza Fatima Nadeem	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/1/2026 11:02	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
419889	Fatima	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	3.75/10	37.50%	CLEAN	10/4/2026 18:04	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Incorrect (Wrong condition: E4>45000. Expected C108>45000.)"
547959	Maryam Naz	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/2/2026 14:51	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
542657	Bushra Hamza	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	9/28/2026 20:40	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
505472	Huzaifa	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/2/2026 5:22	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
485977	Emaan Iftikhar	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	3.75/10	37.50%	CLEAN	10/2/2026 6:49	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Incorrect (Wrong condition: E108>45000. Expected C108>45000.)"
563934	Muhammad Hassan	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	3.75/10	37.50%	CLEAN	9/30/2026 15:48	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Incorrect (Wrong condition: E4>45000. Expected C108>45000.)"
574277	Sadia Arif	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	3.75/10	37.50%	CLEAN	10/2/2026 14:54	"Q1: Correct
Q2: Correct
Q3: Incorrect (Wrong range: D5:D104. Expected D3:D103 or D4:D103.)
Q4: Correct"
546048	Binish khan	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/12/2026 15:10	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
576096	Muhammad Meezan Ali	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	3.75/10	37.50%	CLEAN	10/2/2026 11:02	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Incorrect (Wrong condition: E108>45000. Expected C108>45000.)"
561319	Zainab Nasir	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	3.75/10	37.50%	CLEAN	10/12/2026 16:00	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Incorrect (Wrong condition: E4>45000. Expected C108>45000.)"
581653	Rooma Yousuf	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	2.5/10	25.00%	CLEAN	10/2/2026 10:49	"Q1: Incorrect (Wrong lookup table: A4:E104. Expected A3:E103 or A4:E103.)
Q2: Correct
Q3: Correct
Q4: Incorrect (Wrong condition: E4>45000. Expected C108>45000.)"
564374	Jawad Naeem	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	3.75/10	37.50%	CLEAN	10/4/2026 5:02	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Incorrect (TRUE result ""High"" hona chahiye. | FALSE result ""Low"" hona chahiye.)"
487818	Syed Dilawar Hashmi	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	2.5/10	25.00%	CLEAN	10/3/2026 4:29	"Q1: Incorrect (VLOOKUP match argument missing. Exact match FALSE/0 is required.)
Q2: Correct
Q3: Correct
Q4: Incorrect (Wrong condition: E4>45000. Expected C108>45000.)"
522981	Nimra Usman	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/2/2026 19:12	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
561998	Syed Umar	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	3.75/10	37.50%	CLEAN	9/30/2026 22:22	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Incorrect (Wrong condition: E4>45000. Expected C108>45000.)"
144204	M.Yahya khan	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	0/10	0.00%	CLEAN	9/25/2026 5:31	"Q1: Incorrect (Expected: =VLOOKUP(""E050"", A4:E103, 5, 0). Aapka: E050)
Q2: Incorrect (Expected: =SUMIF(C4:C103, ""IT"", E4:E103). Aapka: =VLOOKUP(C108,A4:E103,5,0))
Q3: Incorrect (Expected: =COUNTIF(C4:C103, ""Karachi""). Aapka: =SUMIF(C4:C103,C100,E4:E103))
Q4: Incorrect (Expected: =IF(E108 > 45000, ""High"", ""Low""). Aapka: =COUNTIF(D4:D103,D89))"
554588	Ayman Habib	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/1/2026 7:38	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
518123	Mehwish Usman	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/2/2026 17:34	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
1999	1999	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/5	100.00%	CLEAN	9/25/2026 22:33	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
1999	1999	Excel Skill 1: Formulas & Basics	0/10	0.00%	CLEAN	9/26/2026 2:01	"Q1: Incorrect (Sheet 'VLOOKUP' not found.)
Q2: Incorrect (Sheet 'VLOOKUP' not found.)
Q3: Incorrect (Sheet 'VLOOKUP' not found.)
Q4: Incorrect (Sheet 'VLOOKUP' not found.)
Q5: Incorrect (Sheet 'SUMIF & COUNTIF' not found.)
Q6: Incorrect (Sheet 'SUMIF & COUNTIF' not found.)
Q7: Incorrect (Sheet 'SUMIF & COUNTIF' not found.)
Q8: Incorrect (Sheet 'SUMIF & COUNTIF' not found.)
Q9: Incorrect (Sheet 'SUMIF & COUNTIF' not found.)
Q10: Incorrect (Sheet 'SUMIF & COUNTIF' not found.)
Q11: Incorrect (Sheet 'LEFT RIGHT MID' not found.)
Q12: Incorrect (Sheet 'LEFT RIGHT MID' not found.)
Q13: Incorrect (Sheet 'LEFT RIGHT MID' not found.)
Q14: Incorrect (Sheet 'LEFT RIGHT MID' not found.)
Q15: Incorrect (Sheet 'LEFT RIGHT MID' not found.)
Q16: Incorrect (Sheet 'LEFT RIGHT MID' not found.)
Q17: Incorrect (Sheet 'Complex Challenge' not found.)
Q18: Incorrect (Sheet 'Complex Challenge' not found.)
Q19: Incorrect (Sheet 'Complex Challenge' not found.)
Q20: Incorrect (Sheet 'Complex Challenge' not found.)"
536400	Uswah Fatima	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	0/10	0.00%	CLEAN	9/25/2026 9:50	"Q1: Incorrect (Expected: =VLOOKUP(""E050"", A4:E103, 5, 0). Aapka: A2)
Q2: Incorrect (Expected: =SUMIF(C4:C103, ""IT"", E4:E103). Aapka: C2,E2)
Q3: Incorrect (Expected: =COUNTIF(C4:C103, ""Karachi""). Aapka: D2)
Q4: Incorrect (Expected: =IF(E108 > 45000, ""High"", ""Low""). Aapka: E24500)"
551806	Samra fatima	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/11/2026 18:19	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
526063	Muhammad umer	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/12/2026 17:17	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
574933	Sahil chohan	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	2.5/10	25.00%	CLEAN	10/10/2026 4:25	"Q1: Correct
Q2: Correct
Q3: Incorrect (Wrong range: C4:C103. Expected D3:D103 or D4:D103.)
Q4: Incorrect (Wrong condition: E108>45000. Expected C108>45000.)"
433187	Muhammad Anas Khan	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	9/30/2026 17:57	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
547401	Hafsa asif	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/3/2026 9:12	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
546363	Fatima Ali	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/11/2026 15:14	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
286030	Wardah Saher	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/4/2026 15:12	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
114781	Ariba	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	2.5/10	25.00%	CLEAN	10/10/2026 17:34	"Q1: Correct
Q2: Incorrect (Wrong criteria range: C2:C103. Expected C3:C103 or C4:C103.)
Q3: Correct
Q4: Incorrect (Wrong condition: E2>45000. Expected C108>45000.)"
512934	Fakeha zubair	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/3/2026 13:12	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
548135	Arshia Mustafa	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	3.75/10	37.50%	CLEAN	10/10/2026 7:31	"Q1: Incorrect (Wrong lookup value: A53. Expected E050.)
Q2: Correct
Q3: Correct
Q4: Correct"
538092	Shahmeen shaikh	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/2/2026 21:35	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
545454	Daniya Ali	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/3/2026 6:01	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
121616	Maria Muhammad Hussain	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	1.25/10	12.50%	CLEAN	9/26/2026 5:26	"Q1: Incorrect (Expected: =VLOOKUP(""E050"", A4:E103, 5, 0). Aapka: =VLOOKUP(A53,A4:E103,5,0))
Q2: Correct
Q3: Incorrect (Expected: =COUNTIF(C4:C103, ""Karachi""). Aapka: =COUNTIF(D4:D103,""Karachi""))
Q4: Incorrect (Expected: =IF(E108 > 45000, ""High"", ""Low""). Aapka: =IF(E111>45000,""High"",""Low""))"
484543	Iqra	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	3.75/10	37.50%	CLEAN	9/26/2026 6:00	"Q1: Correct
Q2: Correct
Q3: Incorrect (Expected: =COUNTIF(C4:C103, ""Karachi""). Aapka: =COUNTIF(D4:D103,""KARACHI""))
Q4: Correct"
570274	Amaim	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/11/2026 5:25	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
33217	Hunain Ali Arshad	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/1/2026 12:37	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
582884	Umna tanvir	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	2.5/10	25.00%	CLEAN	10/12/2026 17:59	"Q1: Correct
Q2: Correct
Q3: Incorrect (Wrong range: C4:C103. Expected D3:D103 or D4:D103.)
Q4: Incorrect (Wrong condition: E108>45000. Expected C108>45000.)"
581208	M. Hasnain	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	3.75/10	37.50%	CLEAN	10/1/2026 16:54	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Incorrect (Wrong condition: E4>45000. Expected C108>45000.)"
573524	Fahad Hassan	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	9/28/2026 11:46	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
556556	Wania Anwer	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	2.5/10	25.00%	CLEAN	10/11/2026 20:29	"Q1: Incorrect (Wrong lookup table: A2:E101. Expected A3:E103 or A4:E103.)
Q2: Correct
Q3: Correct
Q4: Incorrect (Wrong condition: E4>45000. Expected C108>45000.)"
354723	Mahnoor khan	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	3.75/10	37.50%	CLEAN	10/18/2026 5:28	"Q1: Correct
Q2: Correct
Q3: Incorrect (Wrong range: D2:D103. Expected D3:D103 or D4:D103.)
Q4: Correct"
528567	Duaa Saleem	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	3.75/10	37.50%	CLEAN	10/13/2026 18:36	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Incorrect (Wrong condition: E4>45000. Expected C108>45000.)"
557727	Nimra kanwal	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	3.75/10	37.50%	CLEAN	9/30/2026 13:07	"Q1: Correct
Q2: Correct
Q3: Incorrect (Wrong range: D4:D1033. Expected D3:D103 or D4:D103.)
Q4: Correct"
577238	Muhammad Umair	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/12/2026 15:05	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
552604	Aysha	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	0/10	0.00%	CLEAN	9/27/2026 6:59	"Q1: Incorrect (Expected: =VLOOKUP(""E050"", A4:E103, 5, 0). Aapka: =VLOOKUP(A53,A:E,5,0))
Q2: Incorrect (Expected: =SUMIF(C4:C103, ""IT"", E4:E103). Aapka: =SUMIF(C:C,""IT"",E:E))
Q3: Incorrect (Expected: =COUNTIF(C4:C103, ""Karachi""). Aapka: =COUNTIF(D:D,D103))
Q4: Incorrect (Expected: =IF(E108 > 45000, ""High"", ""Low""). Aapka: None)"
552606	Barerah	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/17/2026 16:35	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
424403	Ayan Khan	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	2.5/10	25.00%	CLEAN	10/2/2026 7:09	"Q1: Correct
Q2: Correct
Q3: Incorrect (Wrong range: B4:D103. Expected D3:D103 or D4:D103.)
Q4: Incorrect (Wrong condition: E108>45000. Expected C108>45000.)"
521110	Fahad Khalid	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	3.75/10	37.50%	CLEAN	10/10/2026 9:26	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Incorrect (Wrong condition: E53>45000. Expected C108>45000.)"
527404	Zinnia zainab	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	3.75/10	37.50%	CLEAN	10/3/2026 6:52	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Incorrect (Wrong condition: E4>45000. Expected C108>45000.)"
476531	Zainab iqra	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	3.75/10	37.50%	CLEAN	9/29/2026 12:25	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Incorrect (Wrong condition: E4>45000. Expected C108>45000.)"
550503	Azhan Ahmed	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/11/2026 6:05	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
499266	Zahid Iqbal	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	5.0/10	50.00%	CLEAN	10/1/2026 2:26	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct"
563040	abdul basit	Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula	0/10	0.00%	CLEAN	10/19/2026 20:03	"Q1: Incorrect (Wrong lookup table: A2:E101. Expected A3:E103 or A4:E103.)
Q2: Incorrect (Wrong criteria range: C2:C101. Expected C3:C103 or C4:C103. | Wrong sum range: E2:E101. Expected E3:E103 or E4:E103.)
Q3: Incorrect (Wrong range: D2:D101. Expected D3:D103 or D4:D103.)
Q4: Incorrect (Wrong condition: E2>45000. Expected C108>45000.)"
123	123	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	0.75/10	7.50%	CLEAN	10/9/2026 20:29	"Q1: Incorrect (VLOOKUP required in C108)
Q2: Incorrect (IFERROR required in C109)
Q3: Incorrect (IFERROR required in C110)
Q4: Incorrect (VLOOKUP required in C111)
Q5: Correct
Q6: Incorrect (YEAR required in C113)
Q7: Correct
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Incorrect (TEXT required in C116)
Q10: Incorrect (TEXT required in C117)
Q11: Incorrect (TEXT required in C118)
Q12: Incorrect (Concatenation (& or CONCATENATE) required in C119)
Q13: Incorrect (AND required in C120)
Q14: Incorrect (AND required in C121)
Q15: Incorrect (AND required in C122)
Q16: Incorrect (IF required in C123)
Q17: Correct
Q18: Incorrect (OR required in C125)
Q19: Incorrect (AND required in C126)
Q20: Incorrect (IF required in C127)"
12155	Uswah Fatima	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	3.75/10	37.50%	CLEAN	10/11/2026 16:37	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Correct
Q11: Incorrect (Expected: =TEXT(C112, ""mmmm""))
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Incorrect (Expected: =OR(C4=""HR"", C4=""Finance""))
Q19: Incorrect (Expected: =AND(OR(D4=""Karachi"", D4=""Lahore""), E4>45000))
Q20: Incorrect (IF required in C127)"
557727	Nimra kanwal	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.25/10	42.50%	CLEAN	10/10/2026 20:17	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Incorrect (Expected: =AND(C4=""Sales"", E4>40000, F4>=2023))
Q16: Correct
Q17: Incorrect (OR required in C124)
Q18: Correct
Q19: Correct
Q20: Incorrect (IF required in C127)"
554673	Afifa Noshad	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.25/10	42.50%	CLEAN	10/12/2026 1:32	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Incorrect (Expected: =YEAR(C112) or =YEAR(DATE(F4, G4, H4)))
Q7: Incorrect (Expected: =MONTH(C112) or =MONTH(DATE(F4, G4, H4)))
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
547401	Hafsa asif	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.25/10	42.50%	CLEAN	10/5/2026 16:31	"Q1: Correct
Q2: Correct
Q3: Incorrect (Expected: =IFERROR(E4/I4, 0))
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Incorrect (Expected: =TEXT(DATE(F4, G4, H4), ""mmmm""))
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Incorrect (Expected: =AND(C4=""Sales"", E4>40000, F4>=2023))
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
552096	Areeja shah	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	3.5/10	35.00%	CLEAN	10/12/2026 15:35	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Incorrect (Expected: =YEAR(C112) or =YEAR(DATE(F4, G4, H4)))
Q7: Incorrect (Expected: =MONTH(C112) or =MONTH(DATE(F4, G4, H4)))
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Incorrect (Expected: =TEXT(C112, ""yyyy-mm-dd""))
Q11: Incorrect (Expected: =TEXT(C112, ""mmmm""))
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Incorrect (Expected: =AND(C4=""Sales"", E4>40000, F4>=2023))
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
144204	M.Yahya khan	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	3.75/10	37.50%	CLEAN	10/10/2026 14:07	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Incorrect (Expected: =YEAR(C112) or =YEAR(DATE(F4, G4, H4)))
Q7: Incorrect (Expected: =MONTH(C112) or =MONTH(DATE(F4, G4, H4)))
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Incorrect (Expected: =TEXT(C112, ""yyyy-mm-dd""))
Q11: Incorrect (Expected: =TEXT(C112, ""mmmm""))
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
546048	Binish khan	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.5/10	45.00%	CLEAN	10/3/2026 15:51	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Incorrect (Expected: =VLOOKUP(""E020"", A4:E103, 5, FALSE))
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Incorrect (Expected: =TEXT(DATE(F4, G4, H4), ""yyyy-mm-dd""))
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
354723	Mahnoor khan	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.5/10	45.00%	CLEAN	10/3/2026 17:04	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Incorrect (Expected: =VLOOKUP(""E020"", A4:E103, 5, FALSE))
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Incorrect (Expected: =TEXT(DATE(F4, G4, H4), ""yyyy-mm-dd""))
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
564630	Muhammad Bashaar khan suri	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	5.0/10	50.00%	CLEAN	10/12/2026 18:35	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
550503	Azhan Ahmed	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.5/10	45.00%	CLEAN	10/11/2026 18:47	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Incorrect (Expected: =TEXT(C112, ""mmmm""))
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Incorrect (Expected: =AND(OR(D4=""Karachi"", D4=""Lahore""), E4>45000))
Q20: Correct"
542657	Bushra Hamza	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.25/10	42.50%	CLEAN	10/4/2026 20:58	"Q1: Incorrect (Expected: =VLOOKUP(""E050"", A4:E103, 5, 0))
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Incorrect (Expected: =TEXT(DATE(F4, G4, H4), ""mmmm""))
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Incorrect (Expected: =OR(D4=""Karachi"", D4=""Lahore""))
Q18: Correct
Q19: Correct
Q20: Correct"
547959	Maryam Naz	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	5.0/10	50.00%	CLEAN	10/12/2026 9:19	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
521110	Fahad Khalid	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	5.0/10	50.00%	CLEAN	10/10/2026 9:39	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
564374	Jawad Naeem	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	5.0/10	50.00%	CLEAN	10/11/2026 4:26	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
541501	Esha Ansari	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.75/10	47.50%	CLEAN	10/11/2026 9:53	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Incorrect (IF required in C127)"
580422	Muhammad Imran	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.5/10	45.00%	CLEAN	10/10/2026 9:36	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Incorrect (Expected: =MONTH(C112) or =MONTH(DATE(F4, G4, H4)))
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
574933	Sahil chohan	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.25/10	42.50%	CLEAN	10/12/2026 7:08	"Q1: Correct
Q2: Correct
Q3: Incorrect (Expected: =IFERROR(E4/I4, 0))
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Incorrect (Expected: =OR(C4=""HR"", C4=""Finance""))
Q19: Correct
Q20: Correct"
564453	Duksh Kumar	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.75/10	47.50%	CLEAN	10/10/2026 19:29	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Incorrect (Expected: =TEXT(C112, ""mmmm""))
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
573524	Fahad Hassan	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.75/10	47.50%	CLEAN	10/10/2026 7:09	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Incorrect (IF required in C127)"
576096	Muhammad Meezan Ali	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	3.25/10	32.50%	CLEAN	10/10/2026 9:19	"Q1: Incorrect (Expected: =VLOOKUP(""E050"", A4:E103, 5, 0))
Q2: Incorrect (Expected: =IFERROR(VLOOKUP(""E999"", A4:H103, 5, 0), ""Not Found""))
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Incorrect (Expected: =YEAR(C112) or =YEAR(DATE(F4, G4, H4)))
Q7: Incorrect (Expected: =MONTH(C112) or =MONTH(DATE(F4, G4, H4)))
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Correct
Q11: Incorrect (Expected: =TEXT(C112, ""mmmm""))
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Incorrect (IF required in C127)"
548135	Arshia Mustafa	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	2.75/10	27.50%	CLEAN	10/10/2026 7:32	"Q1: Incorrect (Expected: =VLOOKUP(""E050"", A4:E103, 5, 0))
Q2: Incorrect (Expected: =IFERROR(VLOOKUP(""E999"", A4:H103, 5, 0), ""Not Found""))
Q3: Incorrect (Expected: =IFERROR(E4/I4, 0))
Q4: Incorrect (Expected: =VLOOKUP(""E020"", A4:E103, 5, FALSE))
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Incorrect (Expected: =TEXT(E4, ""$#,##0""))
Q10: Correct
Q11: Incorrect (Expected: =TEXT(C112, ""mmmm""))
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Incorrect (Expected: =OR(C4=""HR"", C4=""Finance""))
Q19: Incorrect (AND required in C126)
Q20: Incorrect (IF required in C127)"
484543	Iqra	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	3.75/10	37.50%	CLEAN	10/10/2026 17:00	"Q1: Correct
Q2: Incorrect (Expected: =IFERROR(VLOOKUP(""E999"", A4:H103, 5, 0), ""Not Found""))
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Correct
Q11: Incorrect (Expected: =TEXT(C112, ""mmmm""))
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Incorrect (Expected: =AND(C4=""Sales"", E4>40000, F4>=2023))
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Incorrect (Expected: =IF(IFERROR(VLOOKUP(""E050"", A4:H103, 5, 0), 0)>50000, ""High"", ""Low""))"
114781	Ariba	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.75/10	47.50%	CLEAN	10/10/2026 17:16	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Incorrect (Expected: =AND(C4=""Sales"", E4>40000, F4>=2023))
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
559189	Hafiza Fatima Nadeem	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.0/10	40.00%	CLEAN	10/10/2026 19:18	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Incorrect (Expected: =YEAR(C112) or =YEAR(DATE(F4, G4, H4)))
Q7: Incorrect (Expected: =MONTH(C112) or =MONTH(DATE(F4, G4, H4)))
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Incorrect (Expected: =OR(D4=""Karachi"", D4=""Lahore""))
Q18: Correct
Q19: Correct
Q20: Correct"
538092	Shahmeen shaikh	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.0/10	40.00%	CLEAN	10/13/2026 11:08	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Incorrect (Expected: =AND(C4=""IT"", E4>50000))
Q14: Correct
Q15: Incorrect (Expected: =AND(C4=""Sales"", E4>40000, F4>=2023))
Q16: Correct
Q17: Correct
Q18: Incorrect (Expected: =OR(C4=""HR"", C4=""Finance""))
Q19: Incorrect (Expected: =AND(OR(D4=""Karachi"", D4=""Lahore""), E4>45000))
Q20: Correct"
577238	Muhammad Umair	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.75/10	47.50%	CLEAN	10/11/2026 5:31	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Incorrect (IF required in C127)"
570274	Amaim	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.75/10	47.50%	CLEAN	10/11/2026 6:53	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Incorrect (IF required in C127)"
505472	Huzaifa	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.0/10	40.00%	CLEAN	10/11/2026 7:18	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Incorrect (Expected: =YEAR(C112) or =YEAR(DATE(F4, G4, H4)))
Q7: Incorrect (Expected: =MONTH(C112) or =MONTH(DATE(F4, G4, H4)))
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Correct
Q11: Incorrect (Expected: =TEXT(C112, ""mmmm""))
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
582660	Shaheer Asif	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	3.75/10	37.50%	CLEAN	10/11/2026 12:01	"Q1: Incorrect (VLOOKUP required in C108)
Q2: Correct
Q3: Correct
Q4: Incorrect (Expected: =VLOOKUP(""E020"", A4:E103, 5, FALSE))
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Incorrect (Expected: =AND(F4=2024, D4=""Karachi""))
Q15: Correct
Q16: Incorrect (IF required in C123)
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Incorrect (IF required in C127)"
546363	Fatima Ali	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	3.75/10	37.50%	CLEAN	10/11/2026 12:55	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Incorrect (Expected: =YEAR(C112) or =YEAR(DATE(F4, G4, H4)))
Q7: Incorrect (Expected: =MONTH(C112) or =MONTH(DATE(F4, G4, H4)))
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Incorrect (Expected: =TEXT(C112, ""yyyy-mm-dd""))
Q11: Incorrect (Expected: =TEXT(C112, ""mmmm""))
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
556556	Wania Anwer	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	5.0/10	50.00%	CLEAN	10/11/2026 14:07	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
551806	Samra fatima	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	5.0/10	50.00%	CLEAN	10/11/2026 14:21	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
554426	Muhammad Zubair	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.0/10	40.00%	CLEAN	10/11/2026 17:07	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Incorrect (Expected: =MONTH(C112) or =MONTH(DATE(F4, G4, H4)))
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Incorrect (Concatenation (& or CONCATENATE) required in C119)
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Incorrect (IF required in C127)"
487818	Syed Dilawar Hashmi	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	3.5/10	35.00%	CLEAN	10/11/2026 18:04	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Incorrect (Expected: =YEAR(C112) or =YEAR(DATE(F4, G4, H4)))
Q7: Incorrect (Expected: =MONTH(C112) or =MONTH(DATE(F4, G4, H4)))
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Incorrect (OR required in C124)
Q18: Incorrect (Expected: =OR(C4=""HR"", C4=""Finance""))
Q19: Correct
Q20: Incorrect (IF required in C127)"
581653	Rooma Yousuf	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	5.0/10	50.00%	CLEAN	10/11/2026 18:52	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
581208	M. Hasnain	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.5/10	45.00%	CLEAN	10/11/2026 19:25	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Incorrect (Expected: =AND(C4=""Sales"", E4>40000, F4>=2023))
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
563934	Muhammad Hassan	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.0/10	40.00%	CLEAN	10/11/2026 19:50	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Incorrect (Expected: =AND(F4=2024, D4=""Karachi""))
Q15: Incorrect (AND required in C122)
Q16: Correct
Q17: Incorrect (Expected: =OR(D4=""Karachi"", D4=""Lahore""))
Q18: Correct
Q19: Correct
Q20: Correct"
512934	Fakeha zubair	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.0/10	40.00%	CLEAN	10/11/2026 19:54	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Incorrect (Expected: =AND(F4=2024, D4=""Karachi""))
Q15: Incorrect (Expected: =AND(C4=""Sales"", E4>40000, F4>=2023))
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Incorrect (IF required in C127)"
582884	Umna tanvir	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	2.5/10	25.00%	CLEAN	10/11/2026 22:46	"Q1: Incorrect (Expected: =VLOOKUP(""E050"", A4:E103, 5, 0))
Q2: Correct
Q3: Incorrect (Expected: =IFERROR(E4/I4, 0))
Q4: Incorrect (Expected: =VLOOKUP(""E020"", A4:E103, 5, FALSE))
Q5: Correct
Q6: Incorrect (Expected: =YEAR(C112) or =YEAR(DATE(F4, G4, H4)))
Q7: Incorrect (Expected: =MONTH(C112) or =MONTH(DATE(F4, G4, H4)))
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Incorrect (Expected: =TEXT(C112, ""yyyy-mm-dd""))
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Incorrect (Expected: =OR(C4=""HR"", C4=""Finance""))
Q19: Incorrect (AND required in C126)
Q20: Incorrect (Expected: =IF(IFERROR(VLOOKUP(""E050"", A4:H103, 5, 0), 0)>50000, ""High"", ""Low""))"
542815	Muhammad Hassam	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.0/10	40.00%	CLEAN	10/12/2026 7:10	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Incorrect (MONTH required in C114)
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Incorrect (Expected: =AND(C4=""Sales"", E4>40000, F4>=2023))
Q16: Incorrect (Expected: =IF(AND(C4=""IT"", E4>60000), ""Senior"", ""Standard""))
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Incorrect (IF required in C127)"
554588	Ayman Habib	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.25/10	42.50%	CLEAN	10/12/2026 8:09	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Incorrect (Expected: =YEAR(C112) or =YEAR(DATE(F4, G4, H4)))
Q7: Incorrect (Expected: =MONTH(C112) or =MONTH(DATE(F4, G4, H4)))
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
574277	Sadia Arif	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	3.75/10	37.50%	CLEAN	10/12/2026 9:07	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Incorrect (Expected: =YEAR(C112) or =YEAR(DATE(F4, G4, H4)))
Q7: Incorrect (Expected: =MONTH(C112) or =MONTH(DATE(F4, G4, H4)))
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Incorrect (Expected: =IF(AND(C4=""IT"", E4>60000), ""Senior"", ""Standard""))
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Incorrect (IF required in C127)"
567003	Kamran Khalid	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.75/10	47.50%	CLEAN	10/17/2026 11:15	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Incorrect (IF required in C127)"
571660	Zainab	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.0/10	40.00%	CLEAN	10/12/2026 10:41	"Q1: Correct
Q2: Correct
Q3: Incorrect (Expected: =IFERROR(E4/I4, 0))
Q4: Correct
Q5: Correct
Q6: Incorrect (Expected: =YEAR(C112) or =YEAR(DATE(F4, G4, H4)))
Q7: Incorrect (Expected: =MONTH(C112) or =MONTH(DATE(F4, G4, H4)))
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
499266	Zahid Iqbal	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.0/10	40.00%	CLEAN	10/12/2026 12:08	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Incorrect (Expected: =TEXT(C112, ""mmmm""))
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Incorrect (Expected: =AND(C4=""Sales"", E4>40000, F4>=2023))
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Incorrect (AND required in C126)
Q20: Incorrect (IF required in C127)"
545454	Daniya Ali	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.5/10	45.00%	CLEAN	10/12/2026 13:40	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Incorrect (Expected: =AND(C4=""Sales"", E4>40000, F4>=2023))
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Incorrect (IF required in C127)"
518123	Mehwish Usman	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	5.0/10	50.00%	CLEAN	10/12/2026 14:13	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
522981	Nimra Usman	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	5.0/10	50.00%	CLEAN	10/12/2026 14:16	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
555402	Alisha Amjad	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.25/10	42.50%	CLEAN	10/12/2026 14:22	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Incorrect (Expected: =AND(C4=""Sales"", E4>40000, F4>=2023))
Q16: Incorrect (Expected: =IF(AND(C4=""IT"", E4>60000), ""Senior"", ""Standard""))
Q17: Correct
Q18: Correct
Q19: Incorrect (Expected: =AND(OR(D4=""Karachi"", D4=""Lahore""), E4>45000))
Q20: Correct"
121616	Maria Muhammad Hussain	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	3.75/10	37.50%	CLEAN	10/12/2026 16:43	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Incorrect (Expected: =YEAR(C112) or =YEAR(DATE(F4, G4, H4)))
Q7: Incorrect (MONTH required in C114)
Q8: Incorrect (DAY required in C115)
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Incorrect (Expected: =OR(C4=""HR"", C4=""Finance""))
Q19: Correct
Q20: Incorrect (Expected: =IF(IFERROR(VLOOKUP(""E050"", A4:H103, 5, 0), 0)>50000, ""High"", ""Low""))"
264944	Noman Haroon	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	5.0/10	50.00%	CLEAN	10/12/2026 16:45	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
33217	Hunain Ali Arshad	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	5.0/10	50.00%	CLEAN	10/12/2026 16:46	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
286030	Wardah Saher	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.75/10	47.50%	CLEAN	10/12/2026 16:49	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
483031	Roma Arshad	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	3.5/10	35.00%	CLEAN	10/12/2026 17:24	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Incorrect (Expected: =YEAR(C112) or =YEAR(DATE(F4, G4, H4)))
Q7: Incorrect (Expected: =MONTH(C112) or =MONTH(DATE(F4, G4, H4)))
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Incorrect (Expected: =TEXT(E4, ""$#,##0""))
Q10: Correct
Q11: Incorrect (Expected: =TEXT(C112, ""mmmm""))
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Incorrect (Expected: =IF(AND(C4=""IT"", E4>60000), ""Senior"", ""Standard""))
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
528567	Duaa Saleem	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.25/10	42.50%	CLEAN	10/12/2026 17:53	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Incorrect (Expected: =YEAR(C112) or =YEAR(DATE(F4, G4, H4)))
Q7: Incorrect (Expected: =MONTH(C112) or =MONTH(DATE(F4, G4, H4)))
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
433187	Muhammad Anas Khan	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	3.75/10	37.50%	CLEAN	10/12/2026 18:49	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Incorrect (Expected: =TEXT(C112, ""mmmm""))
Q12: Incorrect (Expected: =A4 & "" - "" & B4)
Q13: Correct
Q14: Correct
Q15: Incorrect (Expected: =AND(C4=""Sales"", E4>40000, F4>=2023))
Q16: Correct
Q17: Incorrect (Expected: =OR(D4=""Karachi"", D4=""Lahore""))
Q18: Incorrect (Expected: =OR(C4=""HR"", C4=""Finance""))
Q19: Correct
Q20: Correct"
563040	abdul basit	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.0/10	40.00%	CLEAN	10/12/2026 18:50	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Incorrect (DAY required in C115)
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Incorrect (Concatenation (& or CONCATENATE) required in C119)
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Incorrect (Expected: =IF(AND(C4=""IT"", E4>60000), ""Senior"", ""Standard""))
Q17: Correct
Q18: Correct
Q19: Incorrect (Expected: =AND(OR(D4=""Karachi"", D4=""Lahore""), E4>45000))
Q20: Correct"
545050	Hammad Sharif	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	3.75/10	37.50%	CLEAN	10/12/2026 18:51	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Incorrect (Expected: =YEAR(C112) or =YEAR(DATE(F4, G4, H4)))
Q7: Incorrect (Expected: =MONTH(C112) or =MONTH(DATE(F4, G4, H4)))
Q8: Incorrect (DAY required in C115)
Q9: Correct
Q10: Correct
Q11: Incorrect (Expected: =TEXT(C112, ""mmmm""))
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Incorrect (IF required in C127)"
561998	Syed Umar	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.5/10	45.00%	CLEAN	10/17/2026 17:42	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Incorrect (Concatenation (& or CONCATENATE) required in C119)
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Incorrect (IF required in C127)"
6084	SANIA BINT E AAMIR	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.5/10	45.00%	CLEAN	10/12/2026 19:00	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Incorrect (AND required in C126)
Q20: Incorrect (IF required in C127)"
552148	Waleeja waqar	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	3.75/10	37.50%	CLEAN	10/12/2026 19:05	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Incorrect (Expected: =YEAR(C112) or =YEAR(DATE(F4, G4, H4)))
Q7: Incorrect (Expected: =MONTH(C112) or =MONTH(DATE(F4, G4, H4)))
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Correct
Q11: Incorrect (Expected: =TEXT(C112, ""mmmm""))
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Incorrect (Expected: =AND(OR(D4=""Karachi"", D4=""Lahore""), E4>45000))
Q20: Correct"
526063	Muhammad umer	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	5.0/10	50.00%	CLEAN	10/12/2026 19:13	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
538162	Burraq waqar	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	3.25/10	32.50%	CLEAN	10/12/2026 19:52	"Q1: Correct
Q2: Incorrect (Expected: =IFERROR(VLOOKUP(""E999"", A4:H103, 5, 0), ""Not Found""))
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Incorrect (Expected: =YEAR(C112) or =YEAR(DATE(F4, G4, H4)))
Q7: Incorrect (Expected: =MONTH(C112) or =MONTH(DATE(F4, G4, H4)))
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Correct
Q11: Incorrect (Expected: =TEXT(C112, ""mmmm""))
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Incorrect (Expected: =AND(C4=""Sales"", E4>40000, F4>=2023))
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Incorrect (IF required in C127)"
476531	Zainab iqra	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	3.75/10	37.50%	CLEAN	10/12/2026 20:22	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Incorrect (Expected: =YEAR(C112) or =YEAR(DATE(F4, G4, H4)))
Q7: Incorrect (Expected: =MONTH(C112) or =MONTH(DATE(F4, G4, H4)))
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Incorrect (Expected: =AND(F4=2024, D4=""Karachi""))
Q15: Incorrect (Expected: =AND(C4=""Sales"", E4>40000, F4>=2023))
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
537132	Romaisa Noman	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.5/10	45.00%	CLEAN	10/13/2026 14:12	"Q1: Incorrect (Expected: =IFERROR(VLOOKUP(""E999"", A4:H103, 5, 0), ""Not Found""))
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Incorrect (Expected: =AND(C4=""IT"", E4>50000))
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
561319	Zainab Nasir	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.5/10	45.00%	CLEAN	10/13/2026 14:16	"Q1: Correct
Q2: Incorrect (Expected: =IFERROR(VLOOKUP(""E999"", A4:H103, 5, 0), ""Not Found""))
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Incorrect (Expected: =AND(C4=""IT"", E4>50000))
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
571700	Aisha amir	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.0/10	40.00%	CLEAN	10/17/2026 14:55	"Q1: Correct
Q2: Correct
Q3: Incorrect (Expected: =IFERROR(E4/I4, 0))
Q4: Correct
Q5: Correct
Q6: Incorrect (Expected: =YEAR(C112) or =YEAR(DATE(F4, G4, H4)))
Q7: Incorrect (Expected: =MONTH(C112) or =MONTH(DATE(F4, G4, H4)))
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
552580	Sonia Kafeel	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	3.0/10	30.00%	CLEAN	10/17/2026 17:27	"Q1: Incorrect (VLOOKUP required in C108)
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Incorrect (Expected: =YEAR(C112) or =YEAR(DATE(F4, G4, H4)))
Q7: Incorrect (Expected: =MONTH(C112) or =MONTH(DATE(F4, G4, H4)))
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Incorrect (Expected: =TEXT(C112, ""yyyy-mm-dd""))
Q11: Incorrect (Expected: =TEXT(C112, ""mmmm""))
Q12: Correct
Q13: Correct
Q14: Incorrect (Expected: =AND(F4=2024, D4=""Karachi""))
Q15: Incorrect (Expected: =AND(C4=""Sales"", E4>40000, F4>=2023))
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
552606	Barerah	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	3.5/10	35.00%	CLEAN	10/17/2026 23:00	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Incorrect (Expected: =DAY(C112) or =DAY(DATE(F4, G4, H4)))
Q9: Correct
Q10: Correct
Q11: Incorrect (Expected: =TEXT(C112, ""mmmm""))
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Incorrect (Expected: =AND(C4=""Sales"", E4>40000, F4>=2023))
Q16: Incorrect (Expected: =IF(AND(C4=""IT"", E4>60000), ""Senior"", ""Standard""))
Q17: Incorrect (Expected: =OR(D4=""Karachi"", D4=""Lahore""))
Q18: Incorrect (Expected: =OR(C4=""HR"", C4=""Finance""))
Q19: Correct
Q20: Correct"
574067	Mehtab khan	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.0/10	40.00%	CLEAN	10/18/2026 12:07	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Correct
Q7: Correct
Q8: Correct
Q9: Incorrect (Expected: =TEXT(E4, ""$#,##0""))
Q10: Correct
Q11: Incorrect (Expected: =TEXT(C112, ""mmmm""))
Q12: Correct
Q13: Correct
Q14: Correct
Q15: Incorrect (AND required in C122)
Q16: Incorrect (IF required in C123)
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"
419889	Fatima	Excel Skill 6: IFERROR, DATE, TEXT, AND & OR	4.0/10	40.00%	CLEAN	10/18/2026 16:41	"Q1: Correct
Q2: Correct
Q3: Correct
Q4: Correct
Q5: Correct
Q6: Incorrect (YEAR required in C113)
Q7: Incorrect (MONTH required in C114)
Q8: Incorrect (DAY required in C115)
Q9: Correct
Q10: Correct
Q11: Correct
Q12: Correct
Q13: Correct
Q14: Incorrect (Expected: =AND(F4=2024, D4=""Karachi""))
Q15: Correct
Q16: Correct
Q17: Correct
Q18: Correct
Q19: Correct
Q20: Correct"

"""

with app.app_context():
    db.session.execute(db.delete(ExcelSubmission))
    db.session.commit()
    print("✅ excel_submission cleared.")

    assignments = {a.title.strip(): a.id for a in ExcelSkillsAssignment.query.all()}
    
    # Parse lines handling multiline grade_details correctly
    lines = raw_data.strip().split('\n')
    records = []
    current_record = None

    for line in lines:
        if not line.strip():
            continue
        parts = line.split('\t')
        if len(parts) >= 6 and parts[0].strip().isdigit() and len(parts[0].strip()) < 10:
            if current_record:
                records.append(current_record)
            current_record = parts
        else:
            if current_record:
                # append to grade_details
                current_record[-1] += "\n" + line

    if current_record:
        records.append(current_record)

    print(f"Parsed {len(records)} records.")

    count = 0
    for parts in records:
        if len(parts) < 6:
            continue
        sid = parts[0].strip()
        name = parts[1].strip()
        assign_title = parts[2].strip()
        score_str = parts[3].strip()
        pct_str = parts[4].strip()
        status = parts[5].strip()
        submitted_at_str = parts[6].strip() if len(parts) > 6 else ""
        grade_details = parts[7].strip() if len(parts) > 7 else ""

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

        assign_id = assignments.get(assign_title)
        if not assign_id:
            for title, aid in assignments.items():
                if assign_title.lower() in title.lower() or title.lower() in assign_title.lower():
                    assign_id = aid
                    break
        if not assign_id:
            continue

        try:
            score = float(score_str.split('/')[0])
        except Exception:
            score = 0.0

        try:
            pct = float(pct_str.replace('%', '').strip())
        except Exception:
            pct = 0.0

        try:
            sub_at = datetime.strptime(submitted_at_str, "%m/%d/%Y %H:%M")
        except Exception:
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
    total_count = ExcelSubmission.query.count()
    print(f"✅ Successfully inserted {count} Excel submission records!")
    print(f"Total excel submissions in DB: {total_count}")
