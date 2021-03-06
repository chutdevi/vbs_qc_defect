
	 ) MB
	,M_SOURCE MS
	,VM_DEPARTMENT_CLASS VC		
	WHERE 
	MB.PARENT_ITEM_CD = MS.ITEM_CD(+)
	AND MS.SOURCE_CD = VC.COMP_SEC_CD(+) 
	AND VC.PARENT_SEC_CD  IS NOT NULL
      --AND MB.PARENT_ITEM_CD = 'J100-11510'
  GROUP BY 		
      MB.COMP_TYP
    , MB.PARENT_ITEM_CD
    , MB.COMP_ITEM_CD
    , MB.ITEM_NAME
    , MB.COMP_MODEL
    , MB.USE_TOTAL
    , MB.UNIT
    , VC.PARENT_SEC_CD
    , VC.COMP_SEC_CD
    , MB.UPDATE_DATE
 ORDER BY 5