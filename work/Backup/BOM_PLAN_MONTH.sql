SELECT


    A .ITEM_CD AS ITEM_CD,

    NVL (
        SUM (
            CASE
            WHEN A .STATUS = 1
            AND TO_CHAR (A .PDATE, 'YYYY/MM/DD') >= TO_CHAR (TRUNC(ADD_MONTHS(SYSDATE, 0), 'MM') + 0, 'YYYY/MM/DD')
            AND TO_CHAR (A .PDATE, 'YYYY/MM/DD') <= TO_CHAR (LAST_DAY(TRUNC(ADD_MONTHS(SYSDATE, 0), 'MM') + 0), 'YYYY/MM/DD') THEN
                A .QTY
            END
        ),
        0
    ) AS DATE1
FROM
    (
        SELECT
            1 AS STATUS,
            TD.PLANT_CD AS PLANT,
            TD.SOURCE_CD AS LINE_CD,
            VD.SEC_NM AS LINE_NAME,
            TD.ITEM_CD,
            MI.ITEM_NAME,
            TD.ACPT_PLAN_DATE AS PDATE,
            TD.ODR_QTY AS QTY
        FROM
            (SELECT * FROM T_OD WHERE OD_TYP = 2 AND OUTSIDE_TYP = 1 AND NOT(ODR_STS_TYP = 9 AND TOTAL_RCV_QTY = 0)) TD,
            VM_DEPARTMENT VD,
            M_ITEM MI
        WHERE
            TD.ITEM_CD = MI.ITEM_CD (+)
        AND TD.SOURCE_CD = VD.SEC_CD (+)
        AND NOT (TD.ODR_STS_TYP = 9 AND TD.TOTAL_RCV_QTY = 0)
        UNION ALL
            SELECT
                2 AS STATUS,
                TR.PLANT_CD AS PLANT,
                TR.WS_CD AS LINE_CD,
                VD.SEC_NM AS LINE_NAME,
                TR.ITEM_CD AS ITEM_CD,
                MI.ITEM_NAME AS ITEM_NAME,
                TR.OPR_DATE AS PDATE,
                TR.ACPT_QTY AS QTY
            FROM
                T_OPR_RSLT TR,
                M_ITEM MI,
                VM_DEPARTMENT VD
            WHERE
                TR.ITEM_CD = MI.ITEM_CD (+)
            AND TR.WS_CD = VD.SEC_CD (+)
    ) A,
    VM_DEPARTMENT_CLASS VM,
    VM_DEPARTMENT VDD,
    M_PLANT_ITEM MP
WHERE
    A .LINE_CD = VM.COMP_SEC_CD (+)
AND A .LINE_CD = VDD.SEC_CD (+)
AND A .ITEM_CD = MP.ITEM_CD (+)
GROUP BY

    A .ITEM_CD
