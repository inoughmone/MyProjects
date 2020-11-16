SELECT POLICYCERT_NO,max(COLLECTED_DATE) as COLLECTED_DATE
into #bill_max
FROM TBILLTRXN 
WHERE 1=1
--and POLICYCERT_NO in('GZ001000000005' ,'GZ001000000001' )
group by POLICYCERT_NO

--yuejiao
SELECT t1.* 
into #temp_policy
FROM #bill_max t1
inner join tpaymentdets t2 on t1.POLICYCERT_NO=t2.POLICYCERT_NO 
WHERE t1.COLLECTED_DATE<convert(varchar(25),dateadd(MM,-2,getdate()),111)--2月前
--select convert(varchar(25),dateadd(MM,-2,getdate()),111)
and t2.LATEST_VERSION_FG = 'Y'
and t2.BILL_INTERVAL_CD='M'

--nianjiao
insert into #temp_policy
SELECT t1.* 
FROM #bill_max t1
inner join tpaymentdets t2 on t1.POLICYCERT_NO=t2.POLICYCERT_NO 
WHERE t1.COLLECTED_DATE<convert(varchar(25),dateadd(MM,-14,getdate()),111)--14月前
--select convert(varchar(25),dateadd(MM,-14,getdate()),111)
and t2.LATEST_VERSION_FG = 'Y'
and t2.BILL_INTERVAL_CD='A'

--result
SELECT A.POLICYCERT_NO, R.POLICY_EFF_DATE, 
--H.PHOLD_FAMILY_NAME, H.Social_sec_no,H.PHOLD_1_ADDR, H.PHOLD_2_ADDR, h.PHOLD_3_ADDR,H.PHOLD_EMAIL, 
PL.PLAN_NAME, C.SPONSOR_NAME, A.ANN_PREMIUM_AMT, 
--H.phold_home_no, H.phold_mobil_no, 
E.account_no, A.billing_org_cd,E.bank_name,tb.BSB_ADDR as Branch, E.bill_freq_cd, F.BILLING_ORG_NAME
,A.POLICY_HOLDER_NO
,max(t1.COLLECTED_DATE) as latest_COLLECTED_DATE
INTO #tresult1
FROM #temp_policy pol,TPOLICYCERT A, TPOLICYCERTRIDER R, --tpolicycertholder H, 
tbillorg F, tpaymentdets E, tplandets pl, TCAMPDETS C,TBILLTRXN t1,TBANKBRANCH tb
where 1=1
AND pol.POLICYCERT_NO=A.POLICYCERT_NO
AND A.POLICYCERT_NO=R.POLICYCERT_NO
AND R.LATEST_VERSION_FG='Y'
--AND A.POLICY_HOLDER_NO=H.POLICY_HOLDER_NO
AND A.BILLING_ORG_CD=F.BILLING_ORG_CD
--AND F.BILLING_ORG_CD IN ('BJ0003','JS0001','SH0016','ZJ0001','GZ0007')
AND A.POLICYCERT_NO=E.POLICYCERT_NO
AND E.LATEST_VERSION_FG='Y'
--AND E.BSB_NO='000009'
AND R.PLAN_NO=PL.PLAN_NO
AND R.CAMPAIGN_CD=C.CAMPAIGN_CD
AND A.STATUS_CD='A'
--and A.POLICYCERT_NO='GZ050000000017'
and t1.POLICYCERT_NO=a.POLICYCERT_NO
--and t1.COLLECTED_FG='y'
--and r.POLICY_REN_DATE between '2020-08-01' and '2020-08-31'
and e.bsb_no=tb.bsb_no
group BY A.POLICYCERT_NO, R.POLICY_EFF_DATE, 
--H.PHOLD_FAMILY_NAME,H.Social_sec_no,H.PHOLD_1_ADDR, H.PHOLD_2_ADDR, h.PHOLD_3_ADDR,H.PHOLD_EMAIL, 
PL.PLAN_NAME, C.SPONSOR_NAME, A.ANN_PREMIUM_AMT, 
--H.phold_home_no, H.phold_mobil_no, 
E.account_no, A.billing_org_cd,E.bank_name,tb.BSB_ADDR,E.bill_freq_cd, F.BILLING_ORG_NAME
,A.POLICY_HOLDER_NO



SELECT t.POLICYCERT_NO, t.POLICY_EFF_DATE, 
H.PHOLD_FAMILY_NAME, H.Social_sec_no,
H.PHOLD_1_ADDR, H.PHOLD_2_ADDR, h.PHOLD_3_ADDR,H.PHOLD_EMAIL, 
t.PLAN_NAME, t.SPONSOR_NAME, t.ANN_PREMIUM_AMT, 
H.phold_home_no, H.phold_mobil_no, 
t.account_no, 
t.billing_org_cd,t.bank_name,t.Branch, t.bill_freq_cd, t.BILLING_ORG_NAME
,t.latest_COLLECTED_DATE 
FROM #tresult1 t INNER JOIN  tpolicycertholder H
ON t.POLICY_HOLDER_NO = h.POLICY_HOLDER_NO



GO



