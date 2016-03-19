Attribute VB_Name = "Module_Calibrations"
Option Base 1
Public Function loadCalibrations(ByVal nucl As String, ByVal wksheet As String) As Integer
    
    Select Case nucl
        Case Is = "10Be"
            col = "B"
            Call addentry(wksheet, col, 1, "N(" & nucl & ")", "Age", "Lat", "Elev", "Reference")
            Call addentry(wksheet, col, 2, 499463, 13000, 37.982, 3180, "W86-1, Nishiizumi et al. (1989)")
            Call addentry(wksheet, col, 3, 579807, 13000, 37.982, 3180, "W86-3, Nishiizumi et al. (1989)")
            Call addentry(wksheet, col, 4, 695960, 13000, 37.421, 3540, "W86-4, Nishiizumi et al. (1989)")
            Call addentry(wksheet, col, 5, 709546, 13000, 37.419, 3558, "W86-5, Nishiizumi et al. (1989)")
            Call addentry(wksheet, col, 6, 719593, 13000, 37.419, 3558, "W86-6, Nishiizumi et al. (1989)")
            Call addentry(wksheet, col, 7, 707034, 13000, 37.419, 3556, "W86-8, Nishiizumi et al. (1989)")
            Call addentry(wksheet, col, 8, 340883, 13000, 38.87, 2430, "W86-10, Nishiizumi et al. (1989)")
            Call addentry(wksheet, col, 9, 335708, 13000, 38.871, 2452, "W86-11, Nishiizumi et al. (1989)")
            Call addentry(wksheet, col, 10, 371299, 13000, 38.871, 2452, "W86-12, Nishiizumi et al. (1989)")
            Call addentry(wksheet, col, 11, 299290, 13000, 38.855, 2145, "W86-13, Nishiizumi et al. (1989)")
            Call addentry(wksheet, col, 12, 211855, 12189, 47.12, 1675, "KOE4, Kubik et al. (1998)")
            Call addentry(wksheet, col, 13, 198248, 13317, 47.12, 1675, "KOE5, Kubik et al. (1998)")
            Call addentry(wksheet, col, 14, 223570, 13020, 47.12, 1680, "KOE6, Kubik et al. (1998)")
            Call addentry(wksheet, col, 15, 184936, 11321, 47.12, 1645, "KOE20, Kubik et al. (1998)")
            Call addentry(wksheet, col, 16, 211949, 15077, 47.12, 1675, "KOE101, Kubik et al. (1998)")
            Call addentry(wksheet, col, 17, 81800, "13,840", 44.29, 357, "06-NE-010-LIT, Balco et al. (2009)")
            Call addentry(wksheet, col, 18, 80600, "13,840", 44.29, 357, "06-NE-011-LIT, Balco et al. (2009)")
            Call addentry(wksheet, col, 19, 88300, "13,840", 44.31, 414, "06-NE-012-LIT, Balco et al. (2009)")
            Call addentry(wksheet, col, 20, 81000, "13,840", 44.31, 412, "06-NE-013-LIT, Balco et al. (2009)")
            Call addentry(wksheet, col, 21, 105700, "16,750", 42.3, 304, "06-NE-001-HOL, Balco et al. (2009)")
            Call addentry(wksheet, col, 22, 80400, "15,850", 42.5, 135, "06-NE-002-LEV, Balco et al. (2009)")
            Call addentry(wksheet, col, 23, 87100, "15,850", 42.51, 160, "06-NE-003-LEV, Balco et al. (2009)")
            Call addentry(wksheet, col, 24, 91100, "15,850", 42.5, 154, "06-NE-004-LEV, Balco et al. (2009)")
            Call addentry(wksheet, col, 25, 73400, "15,100", 43.01, 180, "06-NE-005-ASH, Balco et al. (2009)")
            Call addentry(wksheet, col, 26, 72300, "15,100", 43.01, 184, "06-NE-006-ASH, Balco et al. (2009)")
            Call addentry(wksheet, col, 27, 33300, "14,590", 43.28, 303, "06-NE-008-PER, Balco et al. (2009)")
            Call addentry(wksheet, col, 28, 79900, "14,590", 43.28, 303, "06-NE-009-PER, Balco et al. (2009)")
            Call addentry(wksheet, col, 29, 64900, "13,180", 44.85, 237, "CH-1, Balco et al. (2009)")
            Call addentry(wksheet, col, 30, 60700, "13,180", 44.85, 236, "CH-2, Balco et al. (2009)")
            Call addentry(wksheet, col, 31, 58700, "13,180", 44.84, 226, "CH-3, Balco et al. (2009)")
            Call addentry(wksheet, col, 32, 66400, "13,180", 44.84, 226, "CH-4, Balco et al. (2009)")
            Call addentry(wksheet, col, 33, 65900, "13,180", 44.84, 226, "CH-5, Balco et al. (2009)")
            Call addentry(wksheet, col, 34, 67700, "13,180", 44.84, 226, "CH-6, Balco et al. (2009)")
            Call addentry(wksheet, col, 35, 70900, "13,180", 44.87, 259, "CH-7, Balco et al. (2009)")
            Call addentry(wksheet, col, 36, 35900, "8,100", 69.84, 65, "CI2-01-01, Balco et al. (2009)")
            Call addentry(wksheet, col, 37, 38300, "8,100", 69.83, 65, "CI2-01-02, Balco et al. (2009)")
            Call addentry(wksheet, col, 38, 40200, "8,100", 69.83, 72, "CR03-90, Balco et al. (2009)")
            Call addentry(wksheet, col, 39, 37000, "8,100", 69.83, 67, "CR03-91, Balco et al. (2009)")
            Call addentry(wksheet, col, 40, 40200, "8,100", 69.83, 67, "CR03-92, Balco et al. (2009)")
            Call addentry(wksheet, col, 41, 41100, "8,100", 69.83, 67, "CR03-93, Balco et al. (2009)")
            Call addentry(wksheet, col, 42, 42300, "8,100", 69.83, 65, "CR03-94, Balco et al. (2009)")
            Call addentry(wksheet, col, 43, 52942, 11650, 59.7808, 77, "YDC08-2, Goehring et al. (2012)")
            Call addentry(wksheet, col, 44, 54994, 11650, 59.78113, 79, "YDC08-3, Goehring et al. (2012)")
            Call addentry(wksheet, col, 45, 54075, 11650, 59.78096, 74, "YDC08-4, Goehring et al. (2012)")
            Call addentry(wksheet, col, 46, 57567, 11650, 59.78158, 76, "YDC08-5, Goehring et al. (2012)")
            Call addentry(wksheet, col, 47, 61760, 11650, 59.81646, 99, "YDC08-7, Goehring et al. (2012)")
            Call addentry(wksheet, col, 48, 56640, 11650, 59.81384, 91, "YDC08-8, Goehring et al. (2012)")
            Call addentry(wksheet, col, 49, 58949, 11650, 59.81395, 87, "YDC08-9, Goehring et al. (2012)")
            Call addentry(wksheet, col, 50, 57230, 11650, 59.81367, 88, "YDC08-10, Goehring et al. (2012)")
            Call addentry(wksheet, col, 51, 32489, 6070, 61.66674, 127, "OL08-1, Goehring et al. (2012)")
            Call addentry(wksheet, col, 52, 30114, 6070, 61.6669, 133, "OL08-3, Goehring et al. (2012)")
            Call addentry(wksheet, col, 53, 28002, 6070, 61.66624, 146, "OL08-5, Goehring et al. (2012)")
            Call addentry(wksheet, col, 54, 31803, 6070, 61.66594, 134, "OL08-7, Goehring et al. (2012)")
            Call addentry(wksheet, col, 55, 30371, 6070, 61.66559, 146, "OL08-9, Goehring et al. (2012)")
            Call addentry(wksheet, col, 56, 38182, 6070, 61.6665, 127, "OL08-11, Goehring et al. (2012)")
            Call addentry(wksheet, col, 57, 30770, 6070, 61.66653, 133, "OL08-13, Goehring et al. (2012)")
            Call addentry(wksheet, col, 58, 88806, 9690, -43.57605, 1024.9, "MR-08-01, Putnam et al. (2010)")
            Call addentry(wksheet, col, 59, 93144, 9690, -43.57581, 1024.5, "MR-08-02, Putnam et al. (2010)")
            Call addentry(wksheet, col, 60, 90235, 9690, -43.57452, 1029.4, "MR-08-03, Putnam et al. (2010)")
            Call addentry(wksheet, col, 61, 90203, 9690, -43.57444, 1028.4, "MR-08-04, Putnam et al. (2010)")
            Call addentry(wksheet, col, 62, 91697, 9690, -43.57435, 1032.1, "MR-08-05, Putnam et al. (2010)")
            Call addentry(wksheet, col, 63, 91685, 9690, -43.57751, 1027.6, "MR-08-13, Putnam et al. (2010)")
            Call addentry(wksheet, col, 64, 90500, 9690, -43.57787, 1032, "MR-08-14, Putnam et al. (2010)")
            Call addentry(wksheet, col, 65, 69394, 12500, -50.0843, 284.79, "HE-06-02, Kaplan et al. (2011)")
            Call addentry(wksheet, col, 66, 65293, 12500, -50.0841, 271.15, "HE-06-04, Kaplan et al. (2011)")
            Call addentry(wksheet, col, 67, 73740, 12500, -50.0826, 340.62, "HE-07-11, Kaplan et al. (2011)")
            Call addentry(wksheet, col, 68, 66010, 12500, -50.083, 288.42, "HE-06-06, Kaplan et al. (2011)")
            Call addentry(wksheet, col, 69, 66641, 12500, -50.0824, 289.58, "HE-06-07, Kaplan et al. (2011)")
            Call addentry(wksheet, col, 70, 73963, 12500, -50.0797, 292.5, "HE-07-12, Kaplan et al. (2011)")
            Call addentry(wksheet, col, 71, 66531, 12500, -50.0796, 292.12, "HE-07-13, Kaplan et al. (2011)")
            Call addentry(wksheet, col, 72, 67432, 12500, -50.0811, 286.35, "HE-07-10, Kaplan et al. (2011)")
            Call addentry(wksheet, col, 73, 65528, 12500, -50.0816, 259.95, "HE-06-05, Kaplan et al. (2011)")
            Call addentry(wksheet, col, 74, 67134, 12500, -50.0821, 284.29, "HE-06-08, Kaplan et al. (2011)")
            Call addentry(wksheet, col, 75, 62268, 12500, -50.0841, 281.66, "HE-06-01, Kaplan et al. (2011)")
            Call addentry(wksheet, col, 76, 65772, 12500, -50.0841, 277.46, "HE-06-03, Kaplan et al. (2011)")
            Call addentry(wksheet, col, 77, 72778, 12500, -50.17011, 240, "EQ-08-01, Kaplan et al. (2011)")
            Call addentry(wksheet, col, 78, 68128, 12500, -50.1774, 252, "EQ-08-06, Kaplan et al. (2011)")
            Call addentry(wksheet, col, 79, 71024, 12500, -50.17729, 238, "EQ-08-05, Kaplan et al. (2011)")
            Call addentry(wksheet, col, 80, 70073, 12500, -50.28208, 221, "PBS-08-09, Kaplan et al. (2011)")
            Call addentry(wksheet, col, 81, 54504, 12500, -50.17293, 245, "EQ-08-04, Kaplan et al. (2011)")
            Call addentry(wksheet, col, 82, 68710, 12500, -50.28575, 216, "PBS-08-04, Kaplan et al. (2011)")
            Call addentry(wksheet, col, 83, 66515, 12500, -50.28425, 218, "PBS-08-06, Kaplan et al. (2011)")
            Call addentry(wksheet, col, 84, 65514, 12500, -50.28002, 215, "PBS-08-11, Kaplan et al. (2011)")
            Call addentry(wksheet, col, 85, 67087, 12500, -50.28659, 213, "PBS-08-02, Kaplan et al. (2011)")
            Call addentry(wksheet, col, 86, 47200, 11424, 68.9112, 71, "040906-03, Fenton et al. (2012)")
            Call addentry(wksheet, col, 87, 49000, 11424, 68.9112, 41, "040906-04, Fenton et al. (2012)")
            Call addentry(wksheet, col, 88, 49500, 11424, 68.9114, 47, "040906-02, Fenton et al. (2012)")
            Call addentry(wksheet, col, 89, 55400, 10942, 69.2135, 109, "060906-14, Fenton et al. (2012)")
            Call addentry(wksheet, col, 90, 52000, 10942, 69.2134, 115, "060906-15, Fenton et al. (2012)")
            Call addentry(wksheet, col, 91, 55700, 10942, 69.2148, 92, "060906-16, Fenton et al. (2012)")
            Call addentry(wksheet, col, 92, 55600, 9245, 69.2844, 350, "11QOO-01, Young et al. (2013)")
            Call addentry(wksheet, col, 93, 55800, 9245, 69.2844, 350, "11QOO-02, Young et al. (2013)")
            Call addentry(wksheet, col, 94, 57100, 9245, 69.2844, 350, "11QOO-03, Young et al. (2013)")
            Call addentry(wksheet, col, 95, 55400, 9245, 69.2844, 350, "11QOO-04, Young et al. (2013)")
            Call addentry(wksheet, col, 96, 56400, 9245, 69.2842, 350, "11QOO-05, Young et al. (2013)")
            Call addentry(wksheet, col, 97, 38900, 8240, 69.2022, 80, "FST08-01, Young et al. (2013)")
            Call addentry(wksheet, col, 98, 36900, 8240, 69.2019, 80, "FST08-02, Young et al. (2013)")
            Call addentry(wksheet, col, 99, 38800, 8240, 69.1131, 175, "09GRO-08, Young et al. (2013)")
            Call addentry(wksheet, col, 100, 52830, 8240, 69.113, 175, "09GRO-09, Young et al. (2013)")
            Call addentry(wksheet, col, 101, 49240, 8240, 69.1129, 175, "09GRO-11, Young et al. (2013)")
            Call addentry(wksheet, col, 102, 51750, 8240, 69.113, 175, "09GRO-12, Young et al. (2013)")
            Call addentry(wksheet, col, 103, 37700, 8250, 69.8353, 65, "CI2-01-1, Young et al. (2013)")
            Call addentry(wksheet, col, 104, 37300, 8250, 69.8345, 65, "CI2-01-2, Young et al. (2013)")
            Call addentry(wksheet, col, 105, 36500, 8250, 69.8302, 72, "CR-03-90, Young et al. (2013)")
            Call addentry(wksheet, col, 106, 36600, 8250, 69.8318, 67, "CR-03-91, Young et al. (2013)")
            Call addentry(wksheet, col, 107, 37200, 8250, 69.8318, 67, "CR-03-92, Young et al. (2013)")
            Call addentry(wksheet, col, 108, 36600, 8250, 69.8324, 67, "CR-03-93, Young et al. (2013)")
            Call addentry(wksheet, col, 109, 39000, 8250, 69.8328, 65, "CR-03-94, Young et al. (2013)")
            Call addentry(wksheet, col, 110, 535000, 12350, -13.945, 4857, "Huancane IIa, Kelly et al. (2015)")
            loadCalibrations = 109
        Case Is = "26Al"
            col = "I"
            Call addentry(wksheet, col, 1, "N(" & nucl & ")", "Age", "Lat", "Elev", "Reference")
            Call addentry(wksheet, col, 2, 3345238, 13000, 37.982, 3180, "W86-1, Nishiizumi et al. (1989)")
            Call addentry(wksheet, col, 3, 3587629, 13000, 37.982, 3180, "W86-3, Nishiizumi et al. (1989)")
            Call addentry(wksheet, col, 4, 4060606, 13000, 37.421, 3540, "W86-4, Nishiizumi et al. (1989)")
            Call addentry(wksheet, col, 5, 4791667, 13000, 37.419, 3558, "W86-5, Nishiizumi et al. (1989)")
            Call addentry(wksheet, col, 6, 4541667, 13000, 37.419, 3558, "W86-6, Nishiizumi et al. (1989)")
            Call addentry(wksheet, col, 7, 4472222, 13000, 37.419, 3556, "W86-8, Nishiizumi et al. (1989)")
            Call addentry(wksheet, col, 8, 2410000, 13000, 38.87, 2430, "W86-10, Nishiizumi et al. (1989)")
            Call addentry(wksheet, col, 9, 2404255, 13000, 38.871, 2452, "W86-11, Nishiizumi et al. (1989)")
            Call addentry(wksheet, col, 10, 2542553, 13000, 38.871, 2452, "W86-12, Nishiizumi et al. (1989)")
            Call addentry(wksheet, col, 11, 2230000, 13000, 38.855, 2145, "W86-13, Nishiizumi et al. (1989)")
            Call addentry(wksheet, col, 12, 4400495, 12000, 43.12, 3231, "WY-92-138, Gosse et al. (1999)")
            Call addentry(wksheet, col, 13, 4179133, 12000, 43.12, 3231, "WY-93-333, Gosse et al. (1999)")
            Call addentry(wksheet, col, 14, 4327813, 12000, 43.12, 3231, "WY-93-334, Gosse et al. (1999)")
            Call addentry(wksheet, col, 15, 3682504, 12000, 43.12, 3231, "WY-93-335, Gosse et al. (1999)")
            Call addentry(wksheet, col, 16, 3768069, 12000, 43.12, 3231, "WY-93-336, Gosse et al. (1999)")
            Call addentry(wksheet, col, 17, 3759505, 12000, 43.12, 3231, "WY-93-339, Gosse et al. (1999)")
            Call addentry(wksheet, col, 18, 898482, 21900, 40.95, 342, "SPA-O-1, Larsen (1996)")
            Call addentry(wksheet, col, 19, 729615, 21900, 40.95, 342, "SPA-O-2, Larsen (1996)")
            Call addentry(wksheet, col, 20, 951157, 21900, 40.95, 342, "SPA-O-3, Larsen (1996)")
            Call addentry(wksheet, col, 21, 850509, 21900, 40.95, 342, "SPA-O-4, Larsen (1996)")
            Call addentry(wksheet, col, 22, 944490, 21900, 40.97, 336, "SPA-O-5, Larsen (1996)")
            Call addentry(wksheet, col, 23, 637824, 21900, 40.97, 336, "SPA-O-6, Larsen (1996)")
            Call addentry(wksheet, col, 24, 773087, 21900, 41, 375, "SWR-B-8, Larsen (1996)")
            Call addentry(wksheet, col, 25, 614613, 21900, 40.95, 253, "SMH-B-9, Larsen (1996)")
            Call addentry(wksheet, col, 26, 886723, 21900, 40.92, 323, "SAF-B-10, Larsen (1996)")
            Call addentry(wksheet, col, 27, 762190, 21900, 40.92, 323, "SAF-B-11, Larsen (1996)")
            Call addentry(wksheet, col, 28, 847850, 21900, 40.92, 323, "SAF-B-12, Larsen (1996)")
            Call addentry(wksheet, col, 29, 858080, 21900, 40.92, 299, "SAF-O-13, Larsen (1996)")
            Call addentry(wksheet, col, 30, 844743, 21900, 40.92, 278, "SAF-O-14, Larsen (1996)")
            loadCalibrations = 29
        Case Is = "21Ne"
            col = "P"
            Call addentry(wksheet, col, 1, "N(" & nucl & ")", "Age", "Lat", "Elev", "Reference")
            Call addentry(wksheet, col, 2, 2678000, 13000, 38, 3340, "W86-8, Niedermann (2000)")
            loadCalibrations = 1
        Case Is = "3He"
            col = "W"
            Call addentry(wksheet, col, 1, "N(" & nucl & ")", "Age", "Lat", "Elev", "Reference")
            Call addentry(wksheet, col, 2, 24300000, 108700, -46.7001, 530, "LBA-98-064, Ackert et al., 2003")
            Call addentry(wksheet, col, 3, 25100000, 108700, -46.7003, 530, "LBA-98-065, Ackert et al., 2003")
            Call addentry(wksheet, col, 4, 21800000, 108700, -46.6199, 400, "LBA-98-093, Ackert et al., 2003")
            Call addentry(wksheet, col, 5, 22000000, 108700, -46.6196, 400, "LBA-98-096, Ackert et al., 2003")
            Call addentry(wksheet, col, 6, 21500000, 108700, -46.5914, 380, "PAT-98-072, Ackert et al., 2003")
            Call addentry(wksheet, col, 7, 19500000, 67800, -47.068, 900, "LBA-01-047, Ackert et al., 2003")
            Call addentry(wksheet, col, 8, 19100000, 67800, -47.068, 900, "LBA-01-047, Ackert et al., 2003")
            Call addentry(wksheet, col, 9, 20500000, 67800, -47.068, 905, "LBA-04-048, Ackert et al., 2003")
            Call addentry(wksheet, col, 10, 18800000, 67800, -47.068, 905, "LBA-01-048, Ackert et al., 2003")
            Call addentry(wksheet, col, 11, 860000, 2453, 44.263, 1561, "10596, Cerling and Craig, 1994")
            Call addentry(wksheet, col, 12, 640000, 2453, 44.263, 1561, "10596, Cerling and Craig, 1994")
            Call addentry(wksheet, col, 13, 970000, 3247, 41.7678, 1335, "10581, Cerling and Craig, 1994")
            Call addentry(wksheet, col, 14, 3130000, 12701, 41.3294, 1198, "10586, Cerling and Craig, 1994")
            Call addentry(wksheet, col, 15, 6250000, 17300, 38.9347, 1455, "8007, Cerling and Craig, 1994")
            Call addentry(wksheet, col, 16, 6460000, 17300, 38.9333, 1455, "9352, Cerling and Craig, 1994")
            Call addentry(wksheet, col, 17, 6310000, 17300, 38.9333, 1455, "9352, Cerling and Craig, 1994")
            Call addentry(wksheet, col, 18, 6440000, 17300, 38.9333, 1455, "9352, Cerling and Craig, 1994")
            Call addentry(wksheet, col, 19, 6490000, 17300, 38.9333, 1455, "9354, Cerling and Craig, 1994")
            Call addentry(wksheet, col, 20, 6190000, 17300, 38.9333, 1455, "9354, Cerling and Craig, 1994")
            Call addentry(wksheet, col, 21, 5930000, 17466, 42.8556, 1380, "9504, Cerling and Craig, 1994")
            Call addentry(wksheet, col, 22, 6200000, 17466, 42.8556, 1380, "9504, Cerling and Craig, 1994")
            Call addentry(wksheet, col, 23, 6250000, 17466, 42.8556, 1380, "9504, Cerling and Craig, 1994")
            Call addentry(wksheet, col, 24, 13700000, 152000, 29.012, 197, "TA1, Dunai and Wijbrans, 2000")
            Call addentry(wksheet, col, 25, 14700000, 152000, 29.012, 197, "TA2, Dunai and Wijbrans, 2000")
            Call addentry(wksheet, col, 26, 14300000, 152000, 29.012, 197, "TA3, Dunai and Wijbrans, 2000")
            Call addentry(wksheet, col, 27, 24100000, 281000, 28.919, 35, "AFB1, Dunai and Wijbrans, 2000")
            Call addentry(wksheet, col, 28, 23200000, 281000, 28.919, 35, "AFB2, Dunai and Wijbrans, 2000")
            Call addentry(wksheet, col, 29, 22300000, 281000, 28.919, 35, "AFB3, Dunai and Wijbrans, 2000")
            Call addentry(wksheet, col, 30, 23900000, 281000, 28.919, 35, "AFB4, Dunai and Wijbrans, 2000")
            Call addentry(wksheet, col, 31, 961000, 2453, 44.29, 1469, "Y1-2799, Licciardi et al., 1999")
            Call addentry(wksheet, col, 32, 967000, 2453, 44.2764, 1530, "Y2-2742, Licciardi et al., 1999")
            Call addentry(wksheet, col, 33, 952000, 2453, 44.2594, 1622, "Y3-2450, Licciardi et al., 1999")
            Call addentry(wksheet, col, 34, 845000, 2453, 44.2709, 1588, "Y4-2833, Licciardi et al., 1999")
            Call addentry(wksheet, col, 35, 820000, 2453, 44.27, 1600, "Y5-2812, Licciardi et al., 1999")
            Call addentry(wksheet, col, 36, 1090000, 2752, 44.2409, 1533, "B1-2676, Licciardi et al., 1999")
            Call addentry(wksheet, col, 37, 785000, 2752, 44.25, 1536, "B2-3018, Licciardi et al., 1999")
            Call addentry(wksheet, col, 38, 851000, 2752, 44.25, 1536, "B2-2674, Licciardi et al., 1999")
            Call addentry(wksheet, col, 39, 1226000, 2752, 44.24, 1536, "B3-3038, Licciardi et al., 1999")
            Call addentry(wksheet, col, 40, 976000, 2752, 44.24, 1536, "B3-3039, Licciardi et al., 1999")
            Call addentry(wksheet, col, 41, 846000, 2752, 44.2406, 1478, "B4-2891, Licciardi et al., 1999")
            Call addentry(wksheet, col, 42, 953000, 2752, 44.2408, 1515, "B5-2776, Licciardi et al., 1999")
            Call addentry(wksheet, col, 43, 667000, 2848, 44.3701, 925, "C1-2098, Licciardi et al., 1999")
            Call addentry(wksheet, col, 44, 715000, 2848, 44.3655, 966, "C2-0914, Licciardi et al., 1999")
            Call addentry(wksheet, col, 45, 644000, 2848, 44.3655, 966, "C2-3004, Licciardi et al., 1999")
            Call addentry(wksheet, col, 46, 577000, 2848, 44.3637, 924, "C3-2830, Licciardi et al., 1999")
            Call addentry(wksheet, col, 47, 674000, 2848, 44.3646, 930, "C4-2906, Licciardi et al., 1999")
            Call addentry(wksheet, col, 48, 696000, 2848, 44.3689, 933, "C5-2877, Licciardi et al., 1999")
            Call addentry(wksheet, col, 49, 2310000, 7091, 43.3763, 1347, "LB1-0886, Licciardi et al., 1999")
            Call addentry(wksheet, col, 50, 2077000, 7091, 43.986, 1216, "LB3-1444, Licciardi et al., 1999")
            Call addentry(wksheet, col, 51, 770000, 4040, 64.3843, 459, "IC02-16-19084, Licciardi et al., 2006")
            Call addentry(wksheet, col, 52, 730000, 4040, 64.3836, 457, "IC02-17-17338, Licciardi et al., 2006")
            Call addentry(wksheet, col, 53, 860000, 4040, 64.3644, 447, "IC02-19-19323, Licciardi et al., 2006")
            Call addentry(wksheet, col, 54, 820000, 4040, 64.3629, 446, "IC02-20-16373, Licciardi et al., 2006")
            Call addentry(wksheet, col, 55, 820000, 5210, 63.9726, 243, "LEIT-1-1067, Licciardi et al., 2006")
            Call addentry(wksheet, col, 56, 930000, 5210, 63.9738, 247, "LEIT-2-1374, Licciardi et al., 2006")
            Call addentry(wksheet, col, 57, 890000, 5210, 63.9765, 273, "LEIT-3-1191, Licciardi et al., 2006")
            Call addentry(wksheet, col, 58, 960000, 5210, 63.9822, 289, "LEIT-4-1131, Licciardi et al., 2006")
            Call addentry(wksheet, col, 59, 1050000, 5210, 63.9839, 277, "LEIT-5-1315, Licciardi et al., 2006")
            Call addentry(wksheet, col, 60, 920000, 5210, 63.9838, 277, "LEIT-5-1129, Licciardi et al., 2006")
            Call addentry(wksheet, col, 61, 1260000, 8060, 64.0592, 96, "BUR-1-2410, Licciardi et al., 2006")
            Call addentry(wksheet, col, 62, 1130000, 8060, 64.0892, 30, "BUR-2-2401, Licciardi et al., 2006")
            Call addentry(wksheet, col, 63, 1260000, 8060, 64.0896, 22, "BUR-3-2713, Licciardi et al., 2006")
            Call addentry(wksheet, col, 64, 1170000, 8060, 64.0869, 22, "BUR-3-2780, Licciardi et al., 2006")
            Call addentry(wksheet, col, 65, 1140000, 8060, 64.0917, 26, "BUR-4-2585, Licciardi et al., 2006")
            Call addentry(wksheet, col, 66, 1100000, 8060, 64.0917, 26, "BUR-4-2854, Licciardi et al., 2006")
            Call addentry(wksheet, col, 67, 990000, 8060, 64.0935, 28, "BUR-5-2732, Licciardi et al., 2006")
            Call addentry(wksheet, col, 68, 1070000, 8060, 64.0871, 27, "BUR-6-2792, Licciardi et al., 2006")
            Call addentry(wksheet, col, 69, 1530000, 10330, 64.1645, 131, "IC02-1-18787, Licciardi et al., 2006")
            Call addentry(wksheet, col, 70, 1170000, 10330, 64.1649, 122, "IC02-7-10543, Licciardi et al., 2006")
            Call addentry(wksheet, col, 71, 1550000, 10330, 64.1562, 121, "IC02-10-25131, Licciardi et al., 2006")
            Call addentry(wksheet, col, 72, 1710000, 10330, 64.1561, 120, "IC02-11-17430, Licciardi et al., 2006")
            Call addentry(wksheet, col, 73, 469912, 550, 19.7, 2327, "KS87-47, Kurz et al., 1990")
            Call addentry(wksheet, col, 74, 90223, 599, 19.1667, 36, "KS87-03, Kurz et al., 1990")
            Call addentry(wksheet, col, 75, 295642, 599, 19.3861, 1985, "KS87-14, Kurz et al., 1990")
            Call addentry(wksheet, col, 76, 256706, 599, 19.3444, 2055, "KS87-15, Kurz et al., 1990")
            Call addentry(wksheet, col, 77, 98279, 599, 19.1778, 42, "T87-4, Kurz et al., 1990")
            Call addentry(wksheet, col, 78, 53436, 599, 19.1778, 42, "T87-4, Kurz et al., 1990")
            Call addentry(wksheet, col, 79, 99890, 729, 19.1444, 18, "KS87-31, Kurz et al., 1990")
            Call addentry(wksheet, col, 80, 152520, 2238, 19.1639, 24, "T87-8, Kurz et al., 1990")
            Call addentry(wksheet, col, 81, 89955, 2238, 19.1639, 24, "T87-8, Kurz et al., 1990")
            Call addentry(wksheet, col, 82, 91566, 2238, 19.175, 36, "KS87-5, Kurz et al., 1990")
            Call addentry(wksheet, col, 83, 93445, 2352, 19.17, 115, "KS87-4, Kurz et al., 1990")
            Call addentry(wksheet, col, 84, 912972, 2770, 19.7, 2303, "KS87-43, Kurz et al., 1990")
            Call addentry(wksheet, col, 85, 335651, 3121, 19.3333, 788, "KS87-13, Kurz et al., 1990")
            Call addentry(wksheet, col, 86, 1388254, 4429, 19.8278, 2370, "KS87-46, Kurz et al., 1990")
            Call addentry(wksheet, col, 87, 1380198, 5468, 19.8139, 2339, "KS87-48, Kurz et al., 1990")
            Call addentry(wksheet, col, 88, 1025751, 5468, 19.8139, 2339, "KS87-48, Kurz et al., 1990")
            Call addentry(wksheet, col, 89, 1895759, 7269, 19.75, 1964, "KS87-42, Kurz et al., 1990")
            Call addentry(wksheet, col, 90, 1154641, 7995, 19.676, 85, "KS87-01C, Kurz et al., 1990")
            Call addentry(wksheet, col, 91, 1066029, 8511, 18.9278, 6, "KS87-08, Kurz et al., 1990")
            Call addentry(wksheet, col, 92, 1259364, 8511, 19.0667, 279, "KS87-07, Kurz et al., 1990")
            Call addentry(wksheet, col, 93, 1090196, 10714, 19.8558, 197, "RM88-9490, Kurz et al., 1990")
            Call addentry(wksheet, col, 94, 6340000, 41000, 37.6058, 190, "SI47, Blard et al., 2006")
            Call addentry(wksheet, col, 95, 6590000, 33000, 37.8485, 820, "SI41, Blard et al., 2006")
            Call addentry(wksheet, col, 96, 720000, 8230, 19.061, 80, "ML1A, Blard et al., 2006")
            Call addentry(wksheet, col, 97, 840000, 8230, 19.054, 40, "ML1B, Blard et al., 2006")
            Call addentry(wksheet, col, 98, 840000, 8230, 19.054, 60, "ML1C, Blard et al., 2006")
            Call addentry(wksheet, col, 99, 240000, 1470, 19.4353, 870, "ML5A, Blard et al., 2006")
            Call addentry(wksheet, col, 100, 27930000, 149000, 19.9887, 840, "MK4, Blard et al., 2006")
            Call addentry(wksheet, col, 101, 6026429, 17300, 38.9333, 1455, "TH-9354-C, Poreda and Cerling, 1992")
            Call addentry(wksheet, col, 102, 5573393, 17300, 38.9333, 1455, "TH-9352-C, Poreda and Cerling, 1992")
            Call addentry(wksheet, col, 103, 5984286, 17300, 38.9333, 1455, "TH-9354-C, Poreda and Cerling, 1992")
            Call addentry(wksheet, col, 104, 6911429, 17300, 38.9333, 1455, "TH-9354-C, Poreda and Cerling, 1992")
            Call addentry(wksheet, col, 105, 15300000, 15255, -19.89035, 3791, "TUN-1, Blard et al. (2013)")
            Call addentry(wksheet, col, 106, 15200000, 15255, -19.88982, 3794, "TUN-2, Blard et al. (2013)")
            Call addentry(wksheet, col, 107, 15900000, 15255, -19.89, 3792, "TUN-3, Blard et al. (2013)")
            Call addentry(wksheet, col, 108, 14600000, 15255, -19.88983, 3794, "TUN-4, Blard et al. (2013)")
            Call addentry(wksheet, col, 109, 15100000, 15255, -19.89087, 3784, "TUN-5, Blard et al. (2013)")
            Call addentry(wksheet, col, 110, 14700000, 15255, -19.88567, 3854, "TU-101, Blard et al. (2013)")
            Call addentry(wksheet, col, 111, 16300000, 15255, -19.88666, 3838, "TU-102, Blard et al. (2013)")
            Call addentry(wksheet, col, 112, 14500000, 15255, -19.88785, 3819, "TU-103, Blard et al. (2013)")
            Call addentry(wksheet, col, 113, 15400000, 15255, -19.88953, 3797, "TU-105, Blard et al. (2013)")
            loadCalibrations = 112
        Case Is = "36Cl"
            col = "AD"
            Call addentry(wksheet, col, 1, "N(" & nucl & ")", "Age", "Lat", "Elev", "Reference")
            Call addentry(wksheet, col, 2, 2883000, 17300, 38.93, 1445, "TH, Stone et al. (1996)")
            loadCalibrations = 1
        Case Is = "14C"
            col = "AK"
            Call addentry(wksheet, col, 1, "N(" & nucl & ")", "Age", "Lat", "Elev", "Reference")
            Call addentry(wksheet, col, 2, 356870, 17400, 41.26, 1594, "PP-4, Miller et al. (2006)")
            loadCalibrations = 1
    End Select
End Function
Private Sub addentry(ByVal wksheet As String, ByVal col As String, row As Integer, ByVal N As Variant, ByVal Age As Variant, ByVal Lat As Variant, ByVal Elev As Variant, ByVal Ref As Variant)
    With Worksheets(wksheet).Range(col & row)
        .Value = N
        .Offset(0, 1).Value = Age
        .Offset(0, 2).Value = Lat
        .Offset(0, 3).Value = Elev
        .Offset(0, 6).Value = Ref
    End With
End Sub