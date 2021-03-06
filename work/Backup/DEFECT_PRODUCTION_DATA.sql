		SELECT
      		  DC.PARENT_SEC_CD PD
		, TR.WS_CD  AS LINE
		, DM.SEC_NM AS LINE_NAME
		, TR.ITEM_CD ITEM_CD 
		, MI.ITEM_NAME ITEM_NAME
		, MP.MODEL 
		, SUM( TR.ACPT_QTY ) QTY
		, TR.ITEM_CD || NULL || TR.WS_CD KEY_CD
		FROM
			T_OPR_RSLT TR
		, VM_DEPARTMENT_CLASS DC
		, VM_DEPARTMENT DM
		, M_PLANT_ITEM MP
		, M_ITEM MI
		WHERE
				TR.WS_CD = DC.COMP_SEC_CD(+)
		AND TR.WS_CD = DM.SEC_CD(+)
		AND TR.ITEM_CD = MI.ITEM_CD(+)
		AND TR.PLANT_CD = MP.PLANT_CD
		AND TR.ITEM_CD = MP.ITEM_CD (+)
		AND	TO_CHAR(TR.OPR_DATE,'YYYY/MM/DD') >= TO_CHAR(TRUNC(ADD_MONTHS(SYSDATE,0),'MM'),'YYYY/MM/DD') AND TO_CHAR(TR.OPR_DATE,'YYYY/MM/DD') <= TO_CHAR(LAST_DAY(ADD_MONTHS(SYSDATE,0)),'YYYY/MM/DD')
		AND TR.ACPT_QTY > 0
		GROUP BY
      		  DC.PARENT_SEC_CD 
		, TR.WS_CD
		, DM.SEC_NM
		, TR.ITEM_CD 
		, MI.ITEM_NAME
		, MP.MODEL
		, TR.ITEM_CD || NULL || TR.WS_CD
		ORDER BY 2