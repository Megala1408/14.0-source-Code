CREATE OR REPLACE PROCEDURE "GET_3M_CODEFINDER_DATA"
    (v_refno IN NUMBER,
v_type IN VARCHAR,v_sep_source IN varchar,
RCT1 IN OUT Globalpkg.RCT1)
AS

----------------------------------------------------------------------------------------------------------------------------
--- Version: 		013
--- Name:		get_3m_codefinder_data
--- Description:	This stored procedure is used by PIMS 3M Codefinder interface (iSGrouper.dll)
--- Created:		14/07/2006 - Based on get_3m_grouping_data stored procedure.
--- Arguments:		1. @v_refno as numeric(9) 	input	Refno from the prof_carer_episodes table
---			2. @v_type as varchar(10)	input	Type of patient episode ('PRCAE')
---			3. v_sep_source as varchar	input	Discharge destination source
---
--- Modification Log
--- Version	By		iASSIST		When		Description
--- 002		AMO				23/11/2006	Modified to work without output parameters.
--- 003		AMO		0608641		29/03/2007	Added extra code to support procedure dates.
--- 004		AMO		MaterHS		21/05/2008	Passing patient age as birth date and admission date.
--- 005		AMO		0629372		21/05/2008	iAssist 0629372.
--- 006		AMO		0629372		19/06/2008	Modified home leave calculation to match an algorithm used in get_3M_grouping_data.iAssist 0629372.
--- 007		AMO		0629372		03/07/2008	Modified ALOS calculation.iAssist 0629372.
--- 008		AMO		0685830		02/02/2010	Removed obsolete  SQL to return procedure dates as this is handled now by IPM client.
--- 009		AMO		0704507		19/11/2010	Increased variable length  for total leave days (v_tld) from varchar(2) to varchar(3).
--- 010		AMO		0717762		04/10/2011	Check for IPM APAC v3.0 billing model.
--- 011		AMO		APAC v4		22/11/2011	APAC v4 upgrade. Added extra input parameter to get correct discharge destination value from DISMT or DISDE domain.
--- 012		AMO		APAC v5		06/03/2012	Corrected billing logic for pre-APAC v3 admissions.
--- 013		AMO		0731196		07/06/2012	Billing logic corrected to support multiple charge accounts linked to the same episode.
---
----------------------------------------------------------------------------------------------------------------------------

v_pasid	VARCHAR(20);
v_surname VARCHAR(100);
v_forename VARCHAR(100);
v_adm_dttm DATE;
v_sep_dttm DATE;
v_birth_dttm DATE;
v_acute_los VARCHAR(10);
v_nalos VARCHAR(10);
v_sex VARCHAR(2);
v_sep_mode VARCHAR(2);
v_mhls VARCHAR(2);
v_tld VARCHAR(3);
v_ilos VARCHAR(2);
v_sdf VARCHAR(2);
v_age_years VARCHAR(2);
v_age_days VARCHAR(10);
--v_hmv VARCHAR(10);
v_prcae_start_dttm DATE;
v_prcae_end_dttm DATE;
v_adm_weight VARCHAR(10);
v_def_ccsxt_code VARCHAR(10);
v_fund_ccsxt_code VARCHAR(10);
v_3m_def_ver_code VARCHAR(10);
v_3m_fund_ver_code VARCHAR(10);
v_diagn_ccsxt	VARCHAR(10);
v_proce_dttm_list VARCHAR(500);

v_area_id		VARCHAR(20);
v_facil_id	VARCHAR(20);
v_patnt_refno		NUMERIC(9);
v_prvsp_refno		NUMERIC(9);
v_sexxx_refno		NUMERIC(9);
v_legsc_refno		NUMERIC(9);
v_dismt_refno		NUMERIC(9);
v_purch_refno		NUMERIC(9);
v_fund_code		VARCHAR(20);
v_inmgt_refno		NUMERIC(9);
v_act_inmgt_refno	NUMERIC(9);
v_prvsn_flag		VARCHAR(2);
v_count			integer;
v_disde_refno		numeric(9);


BEGIN

 BEGIN
   SELECT	prcae.patnt_refno,
		  prcae.prvsp_refno,
		prcae.prcae_refno,		
		  prcae.start_dttm,
		  prcae.end_dttm
   INTO  v_patnt_refno,
			v_prvsp_refno,
			v_prcae_start_dttm,
			v_prcae_end_dttm
	  FROM		PROF_CARER_EPISODES prcae
  	WHERE		prcae.prcae_refno = v_refno;
  EXCEPTION
		WHEN NO_DATA_FOUND THEN
			v_patnt_refno := NULL;
		WHEN OTHERS THEN
			v_patnt_refno := NULL;
 END;

-- Get dates for all existing Procedure codes

	v_proce_dttm_list := NULL;


-- Get Health Fund Code

-- Check if running IPM APAC 3.0 billing model

	BEGIN
	SELECT 1 into v_count from dual where exists (select * from user_objects where object_name='BILL_ACTIVITY_LINKS' and object_type='TABLE');
	EXCEPTION
        WHEN NO_DATA_FOUND THEN
            v_count:=0;
        WHEN OTHERS THEN
           v_count:=0;
	END;
	
	IF v_count=1 THEN

	BEGIN
        execute immediate 'SELECT * from (SELECT purch.main_ident,
                bilat.inmgt_refno
        from bill_activity_links bilal,
            billing_attributes bilat,
            patient_insurance_details patin,
            purchasers purch
        where bilal.sorce_code=''PRVSP''
	and bilal.sorce_refno=:x
	and bilal.archv_flag=''N''
	and bilat.billl_refno=bilal.billl_refno
        and bilat.patin_refno = patin.patin_refno
        and patin.purch_refno = purch.purch_refno
        and bilat.archv_flag=''N''
        and patin.archv_flag=''N''
        and purch.archv_flag=''N''
	order by bilal.create_dttm desc)
	where rownum < 2' into v_fund_code,  v_act_inmgt_refno  using v_prvsp_refno;

        EXCEPTION
        WHEN NO_DATA_FOUND THEN
            v_fund_code := NULL;
            v_act_inmgt_refno:=NULL;
        WHEN OTHERS THEN
            RAISE_APPLICATION_ERROR(-20101, 'Unable to retrieve correct billing attributes.');
	END;

	ELSE

		BEGIN
			SELECT purch.main_ident,
			   bilat.inmgt_refno
			INTO v_fund_code,
			v_act_inmgt_refno
		FROM		PROF_CARER_EPISODES	prcae,
			REFERRALS		refrl,
			BILLS			billl,
			BILLING_ATTRIBUTES	bilat,
			PATIENT_INSURANCE_DETAILS	patin,
			PURCHASERS		purch
		WHERE		prcae.prcae_refno = v_refno
		AND		refrl.refrl_refno = prcae.refrl_refno
		AND		billl.pocar_refno = refrl.pocar_refno
		AND		billl.patnt_refno = v_patnt_refno
		AND		bilat.billl_refno = billl.billl_refno
		AND		patin.patin_refno = bilat.patin_refno
		AND		purch.purch_refno = patin.purch_refno
		AND		NVL(billl.archv_flag,'N')='N'
		AND		NVL(refrl.archv_flag,'N')='N'
		AND		NVL(bilat.archv_flag,'N')='N'
		AND 	ROWNUM<2;
		EXCEPTION
			WHEN NO_DATA_FOUND THEN
			v_fund_code := NULL;
			v_act_inmgt_refno:=NULL;
			WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20101, 'Unable to retrieve correct billing attributes.');
		END;

		IF v_fund_code IS NULL THEN
			BEGIN
			SELECT purch.main_ident,
				bilat.inmgt_refno
			INTO v_fund_code,
				v_act_inmgt_refno
			FROM		PROF_CARER_EPISODES	prcae,
				REFERRALS		refrl,
				BILLS			billl,
				BILLING_ATTRIBUTES	bilat,
				PERIOD_OF_CARE_BILLS	pocbl,
				PATIENT_INSURANCE_DETAILS	patin,
				PURCHASERS	purch
			WHERE		prcae.prcae_refno = v_refno
			AND		refrl.refrl_refno = prcae.refrl_refno
			AND		pocbl.pocar_refno=refrl.pocar_refno
			AND		pocbl.billl_refno=billl.billl_refno
			AND		billl.patnt_refno = v_patnt_refno
			AND		bilat.billl_refno=pocbl.billl_refno
			AND		patin.patin_refno = bilat.patin_refno
			AND		purch.purch_refno = patin.purch_refno
			AND		NVL(billl.archv_flag,'N')='N'
			AND		NVL(refrl.archv_flag,'N')='N'
			AND		NVL(bilat.archv_flag,'N')='N'
			AND 	ROWNUM<2;
			EXCEPTION
			WHEN NO_DATA_FOUND THEN
			v_fund_code := NULL;
			WHEN OTHERS THEN
			RAISE_APPLICATION_ERROR(-20101, 'Unable to retrieve correct billing attributes.');
			END;
		END IF;
	END IF;

--
-- Get required details
  BEGIN
	SELECT patnt.pasid,
			patnt.surname,
			patnt.forename,
			prvsp.admit_dttm,
			prvsp.disch_dttm,
			patnt.sexxx_refno,
			patnt.dttm_of_birth,
			prvsp.dismt_refno,
			prvsp.disde_refno,
			prvsp.admission_weight,
			prvsp.inmgt_refno,
			--prvsp.mech_vent_hrs,
			prvsp.prvsn_end_flag--,
			--TO_NUMBER(TO_CHAR((prvsp.admit_dttm - patnt.dttm_of_birth), 'DD'))
	INTO v_pasid,
			v_surname,
			v_forename,
			v_adm_dttm,
			v_sep_dttm,
			v_sexxx_refno,
			v_birth_dttm,
			v_dismt_refno,
			v_disde_refno,
			v_adm_weight,
			v_inmgt_refno,
			--v_hmv,
			v_prvsn_flag --,
			--v_age_years
	FROM		PROVIDER_SPELLS prvsp,
			PATIENTS patnt
	WHERE		prvsp.prvsp_refno = v_prvsp_refno
	AND		patnt.patnt_refno = prvsp.patnt_refno;
  EXCEPTION
		WHEN NO_DATA_FOUND THEN
			v_pasid := NULL;
		WHEN OTHERS THEN
			v_pasid := NULL;
	END;
--
-- Call the stored procedure that works out the hospitals and funds DRG Versions that are required
--
	Get_3m_Grouper_Versions (v_adm_dttm,v_sep_dttm,v_fund_code,v_def_ccsxt_code,v_fund_ccsxt_code,v_3m_def_ver_code,v_3m_fund_ver_code);
--
--
	Get_3m_Id (v_inmgt_refno,v_sdf);

-- If Version Greater or Equal to 4.1 Then use actual intended management
	IF v_3m_def_ver_code <> '3.1' THEN
		v_sdf := 0;
		IF v_sep_dttm IS NOT NULL THEN
			IF TRUNC(v_adm_dttm) = TRUNC(v_sep_dttm) AND v_prvsn_flag = 'Y' THEN
				v_sdf := 1;
			END IF;
		END IF;
	END IF;
--
	IF v_sep_dttm IS NULL THEN
		v_sep_dttm := v_prcae_end_dttm;
	END IF;
	IF v_sep_dttm IS NULL THEN
		v_sep_dttm := SYSDATE;
	END IF;

-- Calculate Total Leave Days
--
	v_tld := 0;
	select		sum( decode(TRUNC(end_dttm) - TRUNC(start_dttm), 0, 1,
				   	TRUNC(end_dttm) - TRUNC(start_dttm)))
	into		v_tld 
	from		home_leaves
	where		prvsp_refno = v_prvsp_refno
	and		NVL(archv_flag,'N') = 'N'
	and		prvsn_flag = 'N'
	and		end_dttm is not null;
--
-- Calculate Acute LOS
--
	v_acute_los := 0;
	IF v_sep_dttm IS NOT NULL THEN
		v_acute_los := TRUNC(v_sep_dttm - v_adm_dttm) - v_tld; --TO_NUMBER(v_sep_dttm - v_adm_dttm) - v_tld;
	END IF;
--
	v_nalos := 0;
-- Get Mental Health Legal Status at time of admission
--
	Get_3m_Mhls (v_patnt_refno,v_adm_dttm,v_sep_dttm,v_mhls);

-- Get current coding system CCSXT code
  BEGIN
	SELECT odpcd.ccsxt_code
	INTO v_diagn_ccsxt
	FROM ODPCD_CODES odpcd,
		REFERENCE_VALUES ccsxt,
		REFERENCE_VALUES dptyp,
		REFERENCE_VALUE_LINKS rflnk
	WHERE odpcd.ccsxt_code=ccsxt.main_code
	AND ccsxt.rfvdm_code='CCSXT'
	AND rflnk.to_rfval_refno=ccsxt.rfval_refno
	AND rflnk.from_rfval_refno=dptyp.rfval_refno
	AND dptyp.rfvdm_code='DPTYP'
	AND dptyp.main_code='DIAGN'
	AND v_sep_dttm >= odpcd.start_dttm
	AND (v_sep_dttm <= odpcd.end_dttm OR odpcd.end_dttm IS NULL)
	AND odpcd.ccsxt_code NOT IN ('PCCL','MDC')
	AND odpcd.ccsxt_code NOT LIKE 'DRG%'
	AND odpcd.ccsxt_code LIKE 'I10%'
	AND NVL(odpcd.archv_flag,'N')='N'
	AND NVL(ccsxt.archv_flag,'N')='N'
	AND NVL(dptyp.archv_flag,'N')='N'
	AND NVL(rflnk.archv_flag,'N')='N'
	AND rflnk.end_dttm IS NULL
	AND ROWNUM <2;
  EXCEPTION
		WHEN NO_DATA_FOUND THEN
			v_diagn_ccsxt := NULL;
		WHEN OTHERS THEN
			v_diagn_ccsxt := NULL;
	END;

-- Get 3M Alt ID's
	Get_3m_Id (v_sexxx_refno,v_sex);

	if v_sep_source='DISMT' then
	begin
		Get_3m_Id (v_dismt_refno,v_sep_mode);
	end;
	else
	begin
		Get_3m_Id (v_disde_refno,v_sep_mode);
	end;
	end if;

	IF v_sep_mode IS NULL THEN
			v_sep_mode := 9;
	END IF;


	OPEN RCT1 FOR
	SELECT
	Get_3m_Codefinder_Data.v_pasid 			pasid,
	Get_3m_Codefinder_Data.v_surname		patnt_surname,
	Get_3m_Codefinder_Data.v_forename		patnt_forename,
	Get_3m_Codefinder_Data.v_adm_dttm		admit_dttm,
	Get_3m_Codefinder_Data.v_sep_dttm		disch_dttm,
	Get_3m_Codefinder_Data.v_birth_dttm		dob,
	Get_3m_Codefinder_Data.v_acute_los		alos,
	Get_3m_Codefinder_Data.v_nalos			nalos,
	Get_3m_Codefinder_Data.v_sex			sex,
	Get_3m_Codefinder_Data.v_sep_mode		sep_mode,
	Get_3m_Codefinder_Data.v_mhls			mhls,
	Get_3m_Codefinder_Data.v_tld			tld,
	Get_3m_Codefinder_Data.v_ilos			ilos,
	Get_3m_Codefinder_Data.v_sdf			sdf,		-- same day flag
--	GET_3M_CODEFINDER_DATA.v_age_years		age,
	TO_CHAR(Get_3m_Codefinder_Data.v_birth_dttm,'dd/mm/yyyy') || ' ' || TO_CHAR(Get_3m_Codefinder_Data.v_adm_dttm,'dd/mm/yyyy') age,
--	GET_3M_CODEFINDER_DATA.v_age_days		age_in_days,
--	Get_3m_Codefinder_Data.v_hmv			hmv,		-- hours of mechanical ventilation
	Get_3m_Codefinder_Data.v_prcae_start_dttm	prcae_start_dttm,
	Get_3m_Codefinder_Data.v_prcae_end_dttm		prcae_end_dttm,
	Get_3m_Codefinder_Data.v_adm_weight		adm_weight,
	Get_3m_Codefinder_Data.v_def_ccsxt_code		hosp_DRG_ccsxt,
	Get_3m_Codefinder_Data.v_fund_ccsxt_code	fund_DRG_ccsxt,
	Get_3m_Codefinder_Data.v_3m_def_ver_code	ENC_hosp_DRG_ccsxt,
	Get_3m_Codefinder_Data.v_3m_fund_ver_code	ENC_fund_DRG_ccsxt,
	Get_3m_Codefinder_Data.v_diagn_ccsxt		diagn_ccsxt,
	Get_3m_Codefinder_Data.v_proce_dttm_list	proce_dttm_list
	FROM dual;

	EXCEPTION WHEN OTHERS THEN
		  RAISE;
END;
/
GRANT EXECUTE ON GET_3M_CODEFINDER_DATA TO PIMS_USER
/

CREATE OR REPLACE PUBLIC SYNONYM GET_3M_CODEFINDER_DATA FOR GET_3M_CODEFINDER_DATA
/