----BEFORE UPDATE----
SELECT STATUS_CD,POLICYCERT_NO 
INTO #TMP_POLICY
FROM TPOLICYCERTRIDER R
WHERE R.STATUS_CD='E'
AND R.LATEST_VERSION_FG='Y'
AND R.POLICY_EFF_DATE>'2010-1-1'

SELECT * FROM #TMP_POLICY

----BEGIN UPDATE----
UPDATE TPOLICYCERTRIDER SET R.STATUS_CD = 'C'
FROM TPOLICYCERTRIDER R
INNER JOIN #TMP_POLICY P
ON P.POLICYCERT_NO = R.POLICYCERT_NO
AND R.LATEST_VERSION_FG='Y'
AND R.POLICY_EFF_DATE>'2010-1-1'

UPDATE TPOLICYCERT SET PR.STATUS_CD = 'C'
FROM TPOLICYCERT PR
INNER JOIN #TMP_POLICY P
ON P.POLICYCERT_NO = PR.POLICYCERT_NO


----AFTER UPDATE----
SELECT R.STATUS_CD,R.POLICYCERT_NO FROM TPOLICYCERTRIDER R
INNER JOIN #TMP_POLICY P
ON P.POLICYCERT_NO = R.POLICYCERT_NO
AND R.LATEST_VERSION_FG='Y'
AND R.POLICY_EFF_DATE>'2010-1-1'


SELECT R.STATUS_CD,R.POLICYCERT_NO FROM TPOLICYCERT R
INNER JOIN #TMP_POLICY P
ON P.POLICYCERT_NO = R.POLICYCERT_NO

DROP TABLE #TMP_POLICY