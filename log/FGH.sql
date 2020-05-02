 SELECT COUNT(FG.ITEM_CD) CC FROM ( SELECT 
  DG.GURP
, NX.PD
, NX.LINE_CD
, NX.LINE_NAME
, NX.VEND_CD
, NX.VEND_ANAME
, NX.ITEM_CD
, NX.ITEM_NAME
, NX.MODEL
, SUM(NX.RECEIVE) RM_TOTAL
, SUM(NX.ACUT) PD_TOTAL
, SUM(NX.DEFT) NG_TOTAL
, SUM(NX.CD_001) AS CD_001
, SUM(NX.CD_002) AS CD_002
, SUM(NX.CD_003) AS CD_003
, SUM(NX.CD_004) AS CD_004
, SUM(NX.CD_005) AS CD_005
, SUM(NX.CD_006) AS CD_006
, SUM(NX.CD_007) AS CD_007
, SUM(NX.CD_008) AS CD_008
, SUM(NX.CD_009) AS CD_009
, SUM(NX.CD_010) AS CD_010
, SUM(NX.CD_011) AS CD_011
, SUM(NX.CD_012) AS CD_012
, SUM(NX.CD_013) AS CD_013
, SUM(NX.CD_100) AS CD_100
, SUM(NX.CD_101) AS CD_101
, SUM(NX.CD_102) AS CD_102
, SUM(NX.CD_103) AS CD_103
, SUM(NX.CD_104) AS CD_104
, SUM(NX.CD_105) AS CD_105
, SUM(NX.CD_106) AS CD_106
, SUM(NX.CD_107) AS CD_107
, SUM(NX.CD_108) AS CD_108
, SUM(NX.CD_109) AS CD_109
, SUM(NX.CD_110) AS CD_110
, SUM(NX.CD_111) AS CD_111
, SUM(NX.CD_112) AS CD_112
, SUM(NX.CD_113) AS CD_113
, SUM(NX.CD_114) AS CD_114
, SUM(NX.CD_115) AS CD_115
, SUM(NX.CD_117) AS CD_117
, SUM(NX.CD_118) AS CD_118
, SUM(NX.CD_119) AS CD_119
, SUM(NX.CD_120) AS CD_120
, SUM(NX.CD_121) AS CD_121
, SUM(NX.CD_122) AS CD_122
, SUM(NX.CD_123) AS CD_123
, SUM(NX.CD_124) AS CD_124
, SUM(NX.CD_125) AS CD_125
, SUM(NX.CD_126) AS CD_126
, SUM(NX.CD_127) AS CD_127
, SUM(NX.CD_128) AS CD_128
, SUM(NX.CD_129) AS CD_129
, SUM(NX.CD_130) AS CD_130
, SUM(NX.CD_131) AS CD_131
, SUM(NX.CD_132) AS CD_132
, SUM(NX.CD_133) AS CD_133
, SUM(NX.CD_134) AS CD_134
, SUM(NX.CD_135) AS CD_135
, SUM(NX.CD_136) AS CD_136
, SUM(NX.CD_137) AS CD_137
, SUM(NX.CD_201) AS CD_201
, SUM(NX.CD_202) AS CD_202
, SUM(NX.CD_203) AS CD_203
, SUM(NX.CD_204) AS CD_204
, SUM(NX.CD_205) AS CD_205
, SUM(NX.CD_206) AS CD_206
, SUM(NX.CD_207) AS CD_207
, SUM(NX.CD_208) AS CD_208
, SUM(NX.CD_209) AS CD_209
, SUM(NX.CD_210) AS CD_210
, SUM(NX.CD_212) AS CD_212
, SUM(NX.CD_213) AS CD_213
, SUM(NX.CD_214) AS CD_214
, SUM(NX.CD_215) AS CD_215
, SUM(NX.CD_300) AS CD_300
, SUM(NX.CD_301) AS CD_301
, SUM(NX.CD_302) AS CD_302
, SUM(NX.CD_303) AS CD_303
, SUM(NX.CD_304) AS CD_304
, SUM(NX.CD_305) AS CD_305
, SUM(NX.CD_306) AS CD_306
, SUM(NX.CD_307) AS CD_307
, SUM(NX.CD_308) AS CD_308
, SUM(NX.CD_309) AS CD_309
, SUM(NX.CD_310) AS CD_310
, SUM(NX.CD_311) AS CD_311
, SUM(NX.CD_312) AS CD_312
, SUM(NX.CD_313) AS CD_313
, SUM(NX.CD_401) AS CD_401
, SUM(NX.CD_402) AS CD_402
, SUM(NX.CD_403) AS CD_403
, SUM(NX.CD_404) AS CD_404
, SUM(NX.CD_405) AS CD_405
, SUM(NX.CD_406) AS CD_406
, SUM(NX.CD_407) AS CD_407
, SUM(NX.CD_408) AS CD_408
, SUM(NX.CD_409) AS CD_409
, SUM(NX.CD_410) AS CD_410
, SUM(NX.CD_411) AS CD_411
, SUM(NX.CD_412) AS CD_412
, SUM(NX.CD_413) AS CD_413
, SUM(NX.CD_414) AS CD_414
, SUM(NX.CD_500) AS CD_500
, SUM(NX.CD_501) AS CD_501


	FROM

	NG_DATA_EXPL NX

	LEFT OUTER JOIN

	DEFECT_GRP DG

	ON NX.KEY_CK = DG.KEY_CK


	GROUP BY

	DG.GURP
, NX.PD
, NX.LINE_CD
, NX.LINE_NAME
, NX.VEND_CD
, NX.VEND_ANAME
, NX.ITEM_CD
, NX.ITEM_NAME
, NX.MODEL


ORDER BY 2,3,5,6,7 ) FG 
