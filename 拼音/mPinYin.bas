Attribute VB_Name = "mPinYin"
Option Explicit
'------------------------------
'https://github.dev/xiiicode/QuickVB6
'------------------------------
Private Type PINYINMAP
    py As String
    code As Long
End Type

Dim dwPYM(395) As PINYINMAP
Dim bool_init As Boolean
'★┳━━━━━━━━━━━━━━━━━━━━
'┃┃ 2014/11/12 0:51:26 GetPinYin
'┃┃ 获取一个汉字编码的拼音 (如"国",返回"guo")
'┗┻━━━━━━━━━━━━━━━━━━━━
Public Function Quanpin(charcode As Integer) As String
    Dim i As Long
    Dim j As Long
    If bool_init = False Then
        Call PY_Initialize
        bool_init = True
    End If
    If charcode > 0 Then
        Quanpin = Chr(charcode)
        Exit Function
    End If
    If charcode > -10839 Then
        i = 0
    ElseIf charcode > -11868 Then
        i = 26
    ElseIf charcode > -12889 Then
        i = 53
    ElseIf charcode > -13918 Then
        i = 81
    ElseIf charcode > -14942 Then
        i = 124
    ElseIf charcode > -15960 Then
        i = 180
    ElseIf charcode > -16984 Then
        i = 231
    ElseIf charcode > -17998 Then
        i = 269
    ElseIf charcode > -18997 Then
        i = 300
    ElseIf charcode > -19991 Then
        i = 335
    ElseIf charcode > -20320 Then
        i = 381
    End If
    For j = i To 395
        If dwPYM(j).code <= charcode Then
            Quanpin = dwPYM(j).py
            Exit Function
        End If
    Next j
End Function
'★┳━━━━━━━━━━━━━━━━━━━━
'┃┃ 2014/11/12 0:51:26 Get1st
'┃┃ 获取一个汉字编码的拼音首字母 (如"国",返回"G")
'┗┻━━━━━━━━━━━━━━━━━━━━
Public Function GetInitials(charcode As Long) As String
    If charcode > -11056 Then
        Get1st = "Z"
    ElseIf charcode > -11848 Then
        Get1st = "Y"
    ElseIf charcode > -12557 Then
        Get1st = "X"
    ElseIf charcode > -12839 Then
        Get1st = "W"
    ElseIf charcode > -13319 Then
        Get1st = "T"
    ElseIf charcode > -14091 Then
        Get1st = "S"
    ElseIf charcode > -14150 Then
        Get1st = "R"
    ElseIf charcode > -14631 Then
        Get1st = "Q"
    ElseIf charcode > -14915 Then
        Get1st = "P"
    ElseIf charcode > -14923 Then
        Get1st = "O"
    ElseIf charcode > -15166 Then
        Get1st = "N"
    ElseIf charcode > -15641 Then
        Get1st = "M"
    ElseIf charcode > -16213 Then
       Get1st = "L"
    ElseIf charcode > -16475 Then
       Get1st = "K"
    ElseIf charcode > -17418 Then
       Get1st = "J"
    ElseIf charcode > -17923 Then
       Get1st = "H"
    ElseIf charcode > -18240 Then
       Get1st = "G"
    ElseIf charcode > -18527 Then
       Get1st = "F"
    ElseIf charcode > -18711 Then
       Get1st = "E"
    ElseIf charcode > -19219 Then
       Get1st = "D"
    ElseIf charcode > -19776 Then
       Get1st = "C"
    ElseIf charcode > -20284 Then
       Get1st = "B"
    ElseIf charcode > -20320 Then
       Get1st = "A"
    End If
End Function



Private Sub PY_Initialize()
    dwPYM(0).py = "Zuo": dwPYM(0).code = -10254
    dwPYM(1).py = "Zun": dwPYM(1).code = -10256
    dwPYM(2).py = "Zui": dwPYM(2).code = -10260
    dwPYM(3).py = "Zuan": dwPYM(3).code = -10262
    dwPYM(4).py = "Zu": dwPYM(4).code = -10270
    dwPYM(5).py = "Zou": dwPYM(5).code = -10274
    dwPYM(6).py = "Zong": dwPYM(6).code = -10281
    dwPYM(7).py = "Zi": dwPYM(7).code = -10296
    dwPYM(8).py = "Zhuo": dwPYM(8).code = -10307
    dwPYM(9).py = "Zhun": dwPYM(9).code = -10309
    dwPYM(10).py = "Zhui": dwPYM(10).code = -10315
    dwPYM(11).py = "Zhuang": dwPYM(11).code = -10322
    dwPYM(12).py = "Zhuan": dwPYM(12).code = -10328
    dwPYM(13).py = "Zhuai": dwPYM(13).code = -10329
    dwPYM(14).py = "Zhua": dwPYM(14).code = -10331
    dwPYM(15).py = "Zhu": dwPYM(15).code = -10519
    dwPYM(16).py = "Zhou": dwPYM(16).code = -10533
    dwPYM(17).py = "Zhong": dwPYM(17).code = -10544
    dwPYM(18).py = "Zhi": dwPYM(18).code = -10587
    dwPYM(19).py = "Zheng": dwPYM(19).code = -10764
    dwPYM(20).py = "Zhen": dwPYM(20).code = -10780
    dwPYM(21).py = "Zhe": dwPYM(21).code = -10790
    dwPYM(22).py = "Zhao": dwPYM(22).code = -10800
    dwPYM(23).py = "Zhang": dwPYM(23).code = -10815
    dwPYM(24).py = "Zhan": dwPYM(24).code = -10832
    dwPYM(25).py = "Zhai": dwPYM(25).code = -10838
    
    dwPYM(26).py = "Zha": dwPYM(26).code = -11014
    dwPYM(27).py = "Zeng": dwPYM(27).code = -11018
    dwPYM(28).py = "Zen": dwPYM(28).code = -11019
    dwPYM(29).py = "Zei": dwPYM(29).code = -11020
    dwPYM(30).py = "Ze": dwPYM(30).code = -11024
    dwPYM(31).py = "Zao": dwPYM(31).code = -11038
    dwPYM(32).py = "Zang": dwPYM(32).code = -11041
    dwPYM(33).py = "Zan": dwPYM(33).code = -11045
    dwPYM(34).py = "Zai": dwPYM(34).code = -11052
    dwPYM(35).py = "Za": dwPYM(35).code = -11055 '---------------------
    dwPYM(36).py = "Yun": dwPYM(36).code = -11067
    dwPYM(37).py = "Yue": dwPYM(37).code = -11077
    dwPYM(38).py = "Yuan": dwPYM(38).code = -11097
    dwPYM(39).py = "Yu": dwPYM(39).code = -11303
    dwPYM(40).py = "You": dwPYM(40).code = -11324
    dwPYM(41).py = "Yong": dwPYM(41).code = -11339
    dwPYM(42).py = "Yo": dwPYM(42).code = -11340 '哟
    dwPYM(43).py = "Ying": dwPYM(43).code = -11358
    dwPYM(44).py = "Yin": dwPYM(44).code = -11536
    dwPYM(45).py = "Yi": dwPYM(45).code = -11589
    dwPYM(46).py = "Ye": dwPYM(46).code = -11604
    dwPYM(47).py = "Yao": dwPYM(47).code = -11781
    dwPYM(48).py = "Yang": dwPYM(48).code = -11798
    dwPYM(49).py = "Yan": dwPYM(49).code = -11831
    dwPYM(50).py = "Ya": dwPYM(50).code = -11847 '---------------------
    dwPYM(51).py = "Xun": dwPYM(51).code = -11861
    dwPYM(52).py = "Xue": dwPYM(52).code = -11867
    
    dwPYM(53).py = "Xuan": dwPYM(53).code = -12039
    dwPYM(54).py = "Xu": dwPYM(54).code = -12058
    dwPYM(55).py = "Xiu": dwPYM(55).code = -12067
    dwPYM(56).py = "Xiong": dwPYM(56).code = -12074
    dwPYM(57).py = "Xing": dwPYM(57).code = -12089
    dwPYM(58).py = "Xin": dwPYM(58).code = -12099
    dwPYM(59).py = "Xie": dwPYM(59).code = -12120
    dwPYM(60).py = "Xiao": dwPYM(60).code = -12300
    dwPYM(61).py = "Xiang": dwPYM(61).code = -12320
    dwPYM(62).py = "Xian": dwPYM(62).code = -12346
    dwPYM(63).py = "Xia": dwPYM(63).code = -12359
    dwPYM(64).py = "Xi": dwPYM(64).code = -12556 '---------------------
    dwPYM(65).py = "Wu": dwPYM(65).code = -12585
    dwPYM(66).py = "Wo": dwPYM(66).code = -12594
    dwPYM(67).py = "Weng": dwPYM(67).code = -12597
    dwPYM(68).py = "Wen": dwPYM(68).code = -12607
    dwPYM(69).py = "Wei": dwPYM(69).code = -12802
    dwPYM(70).py = "Wang": dwPYM(70).code = -12812
    dwPYM(71).py = "Wan": dwPYM(71).code = -12829
    dwPYM(72).py = "Wai": dwPYM(72).code = -12831
    dwPYM(73).py = "Wa": dwPYM(73).code = -12838 '---------------------
    dwPYM(74).py = "Tuo": dwPYM(74).code = -12849
    dwPYM(75).py = "Tun": dwPYM(75).code = -12852
    dwPYM(76).py = "Tui": dwPYM(76).code = -12858
    dwPYM(77).py = "Tuan": dwPYM(77).code = -12860
    dwPYM(78).py = "Tu": dwPYM(78).code = -12871
    dwPYM(79).py = "Tou": dwPYM(79).code = -12875
    dwPYM(80).py = "Tong": dwPYM(80).code = -12888
    
    dwPYM(81).py = "Ting": dwPYM(81).code = -13060
    dwPYM(82).py = "Tie": dwPYM(82).code = -13063
    dwPYM(83).py = "Tiao": dwPYM(83).code = -13068
    dwPYM(84).py = "Tian": dwPYM(84).code = -13076
    dwPYM(85).py = "Ti": dwPYM(85).code = -13091
    dwPYM(86).py = "Teng": dwPYM(86).code = -13095
    dwPYM(87).py = "Te": dwPYM(87).code = -13096
    dwPYM(88).py = "Tao": dwPYM(88).code = -13107
    dwPYM(89).py = "Tang": dwPYM(89).code = -13120
    dwPYM(90).py = "Tan": dwPYM(90).code = -13138
    dwPYM(91).py = "Tai": dwPYM(91).code = -13147
    dwPYM(92).py = "Ta": dwPYM(92).code = -13318 '---------------------
    dwPYM(93).py = "Suo": dwPYM(93).code = -13326
    dwPYM(94).py = "Sun": dwPYM(94).code = -13329
    dwPYM(95).py = "Sui": dwPYM(95).code = -13340
    dwPYM(96).py = "Suan": dwPYM(96).code = -13343
    dwPYM(97).py = "Su": dwPYM(97).code = -13356
    dwPYM(98).py = "Sou": dwPYM(98).code = -13359
    dwPYM(99).py = "Song": dwPYM(99).code = -13367
    dwPYM(100).py = "Si": dwPYM(100).code = -13383
    dwPYM(101).py = "Shuo": dwPYM(101).code = -13387
    dwPYM(102).py = "Shun": dwPYM(102).code = -13391
    dwPYM(103).py = "Shui": dwPYM(103).code = -13395
    dwPYM(104).py = "Shuang": dwPYM(104).code = -13398
    dwPYM(105).py = "Shuan": dwPYM(105).code = -13400
    dwPYM(106).py = "Shuai": dwPYM(106).code = -13404
    dwPYM(107).py = "Shua": dwPYM(107).code = -13406
    dwPYM(108).py = "Shu": dwPYM(108).code = -13601
    dwPYM(109).py = "Shou": dwPYM(109).code = -13611
    dwPYM(110).py = "Shi": dwPYM(110).code = -13658
    dwPYM(111).py = "Sheng": dwPYM(111).code = -13831
    dwPYM(112).py = "Shen": dwPYM(112).code = -13847
    dwPYM(113).py = "She": dwPYM(113).code = -13859
    dwPYM(114).py = "Shao": dwPYM(114).code = -13870
    dwPYM(115).py = "Shang": dwPYM(115).code = -13878
    dwPYM(116).py = "Shan": dwPYM(116).code = -13894
    dwPYM(117).py = "Shai": dwPYM(117).code = -13896
    dwPYM(118).py = "Sha": dwPYM(118).code = -13905
    dwPYM(119).py = "Seng": dwPYM(119).code = -13906
    dwPYM(120).py = "Sen": dwPYM(120).code = -13907
    dwPYM(121).py = "Se": dwPYM(121).code = -13910
    dwPYM(122).py = "Sao": dwPYM(122).code = -13914
    dwPYM(123).py = "Sang": dwPYM(123).code = -13917
    
    dwPYM(124).py = "San": dwPYM(124).code = -14083
    dwPYM(125).py = "Sai": dwPYM(125).code = -14087
    dwPYM(126).py = "Sa": dwPYM(126).code = -14090 '---------------------
    dwPYM(127).py = "Ruo": dwPYM(127).code = -14092
    dwPYM(128).py = "Run": dwPYM(128).code = -14094
    dwPYM(129).py = "Rui": dwPYM(129).code = -14097
    dwPYM(130).py = "Ruan": dwPYM(130).code = -14099
    dwPYM(131).py = "Ru": dwPYM(131).code = -14109
    dwPYM(132).py = "Rou": dwPYM(132).code = -14112
    dwPYM(133).py = "Rong": dwPYM(133).code = -14122
    dwPYM(134).py = "Ri": dwPYM(134).code = -14123
    dwPYM(135).py = "Reng": dwPYM(135).code = -14125
    dwPYM(136).py = "Ren": dwPYM(136).code = -14135
    dwPYM(137).py = "Re": dwPYM(137).code = -14137
    dwPYM(138).py = "Rao": dwPYM(138).code = -14140
    dwPYM(139).py = "Rang": dwPYM(139).code = -14145
    dwPYM(140).py = "Ran": dwPYM(140).code = -14149 '---------------------
    dwPYM(141).py = "Qun": dwPYM(141).code = -14151
    dwPYM(142).py = "Que": dwPYM(142).code = -14159
    dwPYM(143).py = "Quan": dwPYM(143).code = -14170
    dwPYM(144).py = "Qu": dwPYM(144).code = -14345
    dwPYM(145).py = "Qiu": dwPYM(145).code = -14353
    dwPYM(146).py = "Qiong": dwPYM(146).code = -14355
    dwPYM(147).py = "Qing": dwPYM(147).code = -14368
    dwPYM(148).py = "Qin": dwPYM(148).code = -14379
    dwPYM(149).py = "Qie": dwPYM(149).code = -14384
    dwPYM(150).py = "Qiao": dwPYM(150).code = -14399
    dwPYM(151).py = "Qiang": dwPYM(151).code = -14407
    dwPYM(152).py = "Qian": dwPYM(152).code = -14429
    dwPYM(153).py = "Qia": dwPYM(153).code = -14594
    dwPYM(154).py = "Qi": dwPYM(154).code = -14630 '---------------------
    dwPYM(155).py = "Pu": dwPYM(155).code = -14645
    dwPYM(156).py = "Po": dwPYM(156).code = -14654
    dwPYM(157).py = "Ping": dwPYM(157).code = -14663
    dwPYM(158).py = "Pin": dwPYM(158).code = -14668
    dwPYM(159).py = "Pie": dwPYM(159).code = -14670
    dwPYM(160).py = "Piao": dwPYM(160).code = -14674
    dwPYM(161).py = "Pian": dwPYM(161).code = -14678
    dwPYM(162).py = "Pi": dwPYM(162).code = -14857
    dwPYM(163).py = "Peng": dwPYM(163).code = -14871
    dwPYM(164).py = "Pen": dwPYM(164).code = -14873
    dwPYM(165).py = "Pei": dwPYM(165).code = -14882
    dwPYM(166).py = "Pao": dwPYM(166).code = -14889
    dwPYM(167).py = "Pang": dwPYM(167).code = -14894
    dwPYM(168).py = "Pan": dwPYM(168).code = -14902
    dwPYM(169).py = "Pai": dwPYM(169).code = -14908
    dwPYM(170).py = "Pa": dwPYM(170).code = -14914 '---------------------
    dwPYM(171).py = "Ou": dwPYM(171).code = -14921
    dwPYM(172).py = "O": dwPYM(172).code = -14922 '---------------------
    dwPYM(173).py = "Nuo": dwPYM(173).code = -14926
    dwPYM(174).py = "Nue": dwPYM(174).code = -14928
    dwPYM(175).py = "Nuan": dwPYM(175).code = -14929
    dwPYM(176).py = "Nv": dwPYM(176).code = -14930
    dwPYM(177).py = "Nu": dwPYM(177).code = -14933
    dwPYM(178).py = "Nong": dwPYM(178).code = -14937
    dwPYM(179).py = "Niu": dwPYM(179).code = -14941
    
    dwPYM(180).py = "Ning": dwPYM(180).code = -15109
    dwPYM(181).py = "Nin": dwPYM(181).code = -15110
    dwPYM(182).py = "Nie": dwPYM(182).code = -15117
    dwPYM(183).py = "Niao": dwPYM(183).code = -15119
    dwPYM(184).py = "Niang": dwPYM(184).code = -15121
    dwPYM(185).py = "Nian": dwPYM(185).code = -15128
    dwPYM(186).py = "Ni": dwPYM(186).code = -15139
    dwPYM(187).py = "Neng": dwPYM(187).code = -15140
    dwPYM(188).py = "Nen": dwPYM(188).code = -15141
    dwPYM(189).py = "Nei": dwPYM(189).code = -15143
    dwPYM(190).py = "Ne": dwPYM(190).code = -15144
    dwPYM(191).py = "Nao": dwPYM(191).code = -15149
    dwPYM(192).py = "Nang": dwPYM(192).code = -15150
    dwPYM(193).py = "Nan": dwPYM(193).code = -15153
    dwPYM(194).py = "Nai": dwPYM(194).code = -15158
    dwPYM(195).py = "Na": dwPYM(195).code = -15165 '---------------------
    dwPYM(196).py = "Mu": dwPYM(196).code = -15180
    dwPYM(197).py = "Mou": dwPYM(197).code = -15183
    dwPYM(198).py = "Mo": dwPYM(198).code = -15362
    dwPYM(199).py = "Miu": dwPYM(199).code = -15363
    dwPYM(200).py = "Ming": dwPYM(200).code = -15369
    dwPYM(201).py = "Min": dwPYM(201).code = -15375
    dwPYM(202).py = "Mie": dwPYM(202).code = -15377
    dwPYM(203).py = "Miao": dwPYM(203).code = -15385
    dwPYM(204).py = "Mian": dwPYM(204).code = -15394
    dwPYM(205).py = "Mi": dwPYM(205).code = -15408
    dwPYM(206).py = "Meng": dwPYM(206).code = -15416
    dwPYM(207).py = "Men": dwPYM(207).code = -15419
    dwPYM(208).py = "Mei": dwPYM(208).code = -15435
    dwPYM(209).py = "Me": dwPYM(209).code = -15436
    dwPYM(210).py = "Mao": dwPYM(210).code = -15448
    dwPYM(211).py = "Mang": dwPYM(211).code = -15454
    dwPYM(212).py = "Man": dwPYM(212).code = -15625
    dwPYM(213).py = "Mai": dwPYM(213).code = -15631
    dwPYM(214).py = "Ma": dwPYM(214).code = -15640 '---------------------
    dwPYM(215).py = "Luo": dwPYM(215).code = -15652
    dwPYM(216).py = "Lun": dwPYM(216).code = -15659
    dwPYM(217).py = "Lue": dwPYM(217).code = -15661
    dwPYM(218).py = "Luan": dwPYM(218).code = -15667
    dwPYM(219).py = "Lv": dwPYM(219).code = -15681
    dwPYM(220).py = "Lu": dwPYM(220).code = -15701
    dwPYM(221).py = "Lou": dwPYM(221).code = -15707
    dwPYM(222).py = "Long": dwPYM(222).code = -15878
    dwPYM(223).py = "Liu": dwPYM(223).code = -15889
    dwPYM(224).py = "Ling": dwPYM(224).code = -15903
    dwPYM(225).py = "Lin": dwPYM(225).code = -15915
    dwPYM(226).py = "Lie": dwPYM(226).code = -15920
    dwPYM(227).py = "Liao": dwPYM(227).code = -15933
    dwPYM(228).py = "Liang": dwPYM(228).code = -15944
    dwPYM(229).py = "Lian": dwPYM(229).code = -15958
    dwPYM(230).py = "Lia": dwPYM(230).code = -15959
    
    dwPYM(231).py = "Li": dwPYM(231).code = -16155
    dwPYM(232).py = "Leng": dwPYM(232).code = -16158
    dwPYM(233).py = "Lei": dwPYM(233).code = -16169
    dwPYM(234).py = "Le": dwPYM(234).code = -16171
    dwPYM(235).py = "Lao": dwPYM(235).code = -16180
    dwPYM(236).py = "Lang": dwPYM(236).code = -16187
    dwPYM(237).py = "Lan": dwPYM(237).code = -16202
    dwPYM(238).py = "Lai": dwPYM(238).code = -16205
    dwPYM(239).py = "La": dwPYM(239).code = -16212 '---------------------
    dwPYM(240).py = "Kuo": dwPYM(240).code = -16216
    dwPYM(241).py = "Kun": dwPYM(241).code = -16220
    dwPYM(242).py = "Kui": dwPYM(242).code = -16393
    dwPYM(243).py = "Kuang": dwPYM(243).code = -16401
    dwPYM(244).py = "Kuan": dwPYM(244).code = -16403
    dwPYM(245).py = "Kuai": dwPYM(245).code = -16407
    dwPYM(246).py = "Kua": dwPYM(246).code = -16412
    dwPYM(247).py = "Ku": dwPYM(247).code = -16419
    dwPYM(248).py = "Kou": dwPYM(248).code = -16423
    dwPYM(249).py = "Kong": dwPYM(249).code = -16427
    dwPYM(250).py = "Keng": dwPYM(250).code = -16429
    dwPYM(251).py = "Ken": dwPYM(251).code = -16433
    dwPYM(252).py = "Ke": dwPYM(252).code = -16448
    dwPYM(253).py = "Kao": dwPYM(253).code = -16452
    dwPYM(254).py = "Kang": dwPYM(254).code = -16459
    dwPYM(255).py = "Kan": dwPYM(255).code = -16465
    dwPYM(256).py = "Kai": dwPYM(256).code = -16470
    dwPYM(257).py = "Ka": dwPYM(257).code = -16474 '---------------------
    dwPYM(258).py = "Jun": dwPYM(258).code = -16647
    dwPYM(259).py = "Jue": dwPYM(259).code = -16657
    dwPYM(260).py = "Juan": dwPYM(260).code = -16664
    dwPYM(261).py = "Ju": dwPYM(261).code = -16689
    dwPYM(262).py = "Jiu": dwPYM(262).code = -16706
    dwPYM(263).py = "Jiong": dwPYM(263).code = -16708
    dwPYM(264).py = "Jing": dwPYM(264).code = -16733
    dwPYM(265).py = "Jin": dwPYM(265).code = -16915
    dwPYM(266).py = "Jie": dwPYM(266).code = -16942
    dwPYM(267).py = "Jiao": dwPYM(267).code = -16970
    dwPYM(268).py = "Jiang": dwPYM(268).code = -16983
    
    dwPYM(269).py = "Jian": dwPYM(269).code = -17185
    dwPYM(270).py = "Jia": dwPYM(270).code = -17202
    dwPYM(271).py = "Ji": dwPYM(271).code = -17417 '---------------------
    dwPYM(272).py = "Huo": dwPYM(272).code = -17427
    dwPYM(273).py = "Hun": dwPYM(273).code = -17433
    dwPYM(274).py = "Hui": dwPYM(274).code = -17454
    dwPYM(275).py = "Huang": dwPYM(275).code = -17468
    dwPYM(276).py = "Huan": dwPYM(276).code = -17482
    dwPYM(277).py = "Huai": dwPYM(277).code = -17487
    dwPYM(278).py = "Hua": dwPYM(278).code = -17496
    dwPYM(279).py = "Hu": dwPYM(279).code = -17676
    dwPYM(280).py = "Hou": dwPYM(280).code = -17683
    dwPYM(281).py = "Hong": dwPYM(281).code = -17692
    dwPYM(282).py = "Heng": dwPYM(282).code = -17697
    dwPYM(283).py = "Hen": dwPYM(283).code = -17701
    dwPYM(284).py = "Hei": dwPYM(284).code = -17703
    dwPYM(285).py = "He": dwPYM(285).code = -17721
    dwPYM(286).py = "Hao": dwPYM(286).code = -17730
    dwPYM(287).py = "Hang": dwPYM(287).code = -17733
    dwPYM(288).py = "Han": dwPYM(288).code = -17752
    dwPYM(289).py = "Hai": dwPYM(289).code = -17759
    dwPYM(290).py = "Ha": dwPYM(290).code = -17922 '---------------------
    dwPYM(291).py = "Guo": dwPYM(291).code = -17928
    dwPYM(292).py = "Gun": dwPYM(292).code = -17931
    dwPYM(293).py = "Gui": dwPYM(293).code = -17947
    dwPYM(294).py = "Guang": dwPYM(294).code = -17950
    dwPYM(295).py = "Guan": dwPYM(295).code = -17961
    dwPYM(296).py = "Guai": dwPYM(296).code = -17964
    dwPYM(297).py = "Gua": dwPYM(297).code = -17970
    dwPYM(298).py = "Gu": dwPYM(298).code = -17988
    dwPYM(299).py = "Gou": dwPYM(299).code = -17997
    
    dwPYM(300).py = "Gong": dwPYM(300).code = -18012
    dwPYM(301).py = "Geng": dwPYM(301).code = -18181
    dwPYM(302).py = "Gen": dwPYM(302).code = -18183
    dwPYM(303).py = "Gei": dwPYM(303).code = -18184
    dwPYM(304).py = "Ge": dwPYM(304).code = -18201
    dwPYM(305).py = "Gao": dwPYM(305).code = -18211
    dwPYM(306).py = "Gang": dwPYM(306).code = -18220
    dwPYM(307).py = "Gan": dwPYM(307).code = -18231
    dwPYM(308).py = "Gai": dwPYM(308).code = -18237
    dwPYM(309).py = "Ga": dwPYM(309).code = -18239 '---------------------
    dwPYM(310).py = "Fu": dwPYM(310).code = -18446
    dwPYM(311).py = "Fou": dwPYM(311).code = -18447
    dwPYM(312).py = "Fo": dwPYM(312).code = -18448
    dwPYM(313).py = "Feng": dwPYM(313).code = -18463
    dwPYM(314).py = "Fen": dwPYM(314).code = -18478
    dwPYM(315).py = "Fei": dwPYM(315).code = -18490
    dwPYM(316).py = "Fang": dwPYM(316).code = -18501
    dwPYM(317).py = "Fan": dwPYM(317).code = -18518
    dwPYM(318).py = "Fa": dwPYM(318).code = -18526 '---------------------
    dwPYM(319).py = "Er": dwPYM(319).code = -18696
    dwPYM(320).py = "En": dwPYM(320).code = -18697
    dwPYM(321).py = "E": dwPYM(321).code = -18710 '---------------------
    dwPYM(322).py = "Duo": dwPYM(322).code = -18722
    dwPYM(323).py = "Dun": dwPYM(323).code = -18731
    dwPYM(324).py = "Dui": dwPYM(324).code = -18735
    dwPYM(325).py = "Duan": dwPYM(325).code = -18741
    dwPYM(326).py = "Du": dwPYM(326).code = -18756
    dwPYM(327).py = "Dou": dwPYM(327).code = -18763
    dwPYM(328).py = "Dong": dwPYM(328).code = -18773
    dwPYM(329).py = "Diu": dwPYM(329).code = -18774
    dwPYM(330).py = "Ding": dwPYM(330).code = -18783
    dwPYM(331).py = "Die": dwPYM(331).code = -18952
    dwPYM(332).py = "Diao": dwPYM(332).code = -18961
    dwPYM(333).py = "Dian": dwPYM(333).code = -18977
    dwPYM(334).py = "Di": dwPYM(334).code = -18996
    
    dwPYM(335).py = "Deng": dwPYM(335).code = -19003
    dwPYM(336).py = "De": dwPYM(336).code = -19006
    dwPYM(337).py = "Dao": dwPYM(337).code = -19018
    dwPYM(338).py = "Dang": dwPYM(338).code = -19023
    dwPYM(339).py = "Dan": dwPYM(339).code = -19038
    dwPYM(340).py = "Dai": dwPYM(340).code = -19212
    dwPYM(341).py = "Da": dwPYM(341).code = -19218 '---------------------
    dwPYM(342).py = "Cuo": dwPYM(342).code = -19224
    dwPYM(343).py = "Cun": dwPYM(343).code = -19227
    dwPYM(344).py = "Cui": dwPYM(344).code = -19235
    dwPYM(345).py = "Cuan": dwPYM(345).code = -19238
    dwPYM(346).py = "Cu": dwPYM(346).code = -19242
    dwPYM(347).py = "Cou": dwPYM(347).code = -19243
    dwPYM(348).py = "Cong": dwPYM(348).code = -19249
    dwPYM(349).py = "Ci": dwPYM(349).code = -19261
    dwPYM(350).py = "Chuo": dwPYM(350).code = -19263
    dwPYM(351).py = "Chun": dwPYM(351).code = -19270
    dwPYM(352).py = "Chui": dwPYM(352).code = -19275
    dwPYM(353).py = "Chuang": dwPYM(353).code = -19281
    dwPYM(354).py = "Chuan": dwPYM(354).code = -19288
    dwPYM(355).py = "Chuai": dwPYM(355).code = -19289
    dwPYM(356).py = "Chu": dwPYM(356).code = -19467
    dwPYM(357).py = "Chou": dwPYM(357).code = -19479
    dwPYM(358).py = "Chong": dwPYM(358).code = -19484
    dwPYM(359).py = "Chi": dwPYM(359).code = -19500
    dwPYM(360).py = "Cheng": dwPYM(360).code = -19515
    dwPYM(361).py = "Chen": dwPYM(361).code = -19525
    dwPYM(362).py = "Che": dwPYM(362).code = -19531
    dwPYM(363).py = "Chao": dwPYM(363).code = -19540
    dwPYM(364).py = "Chang": dwPYM(364).code = -19715
    dwPYM(365).py = "Chan": dwPYM(365).code = -19725
    dwPYM(366).py = "Chai": dwPYM(366).code = -19728
    dwPYM(367).py = "Cha": dwPYM(367).code = -19739
    dwPYM(368).py = "Ceng": dwPYM(368).code = -19741
    dwPYM(369).py = "Ce": dwPYM(369).code = -19746
    dwPYM(370).py = "Cao": dwPYM(370).code = -19751
    dwPYM(371).py = "Cang": dwPYM(371).code = -19756
    dwPYM(372).py = "Can": dwPYM(372).code = -19763
    dwPYM(373).py = "Cai": dwPYM(373).code = -19774
    dwPYM(374).py = "Ca": dwPYM(374).code = -19775 '---------------------
    dwPYM(375).py = "Bu": dwPYM(375).code = -19784
    dwPYM(376).py = "Bo": dwPYM(376).code = -19805
    dwPYM(377).py = "Bing": dwPYM(377).code = -19976
    dwPYM(378).py = "Bin": dwPYM(378).code = -19982
    dwPYM(379).py = "Bie": dwPYM(379).code = -19986
    dwPYM(380).py = "Biao": dwPYM(380).code = -19990
    
    dwPYM(381).py = "Bian": dwPYM(381).code = -20002
    dwPYM(382).py = "Bi": dwPYM(382).code = -20026
    dwPYM(383).py = "Beng": dwPYM(383).code = -20032
    dwPYM(384).py = "Ben": dwPYM(384).code = -20036
    dwPYM(385).py = "Bei": dwPYM(385).code = -20051
    dwPYM(386).py = "Bao": dwPYM(386).code = -20230
    dwPYM(387).py = "Bang": dwPYM(387).code = -20242
    dwPYM(388).py = "Ban": dwPYM(388).code = -20257
    dwPYM(389).py = "Bai": dwPYM(389).code = -20265
    dwPYM(390).py = "Ba": dwPYM(390).code = -20283 '---------------------
    dwPYM(391).py = "Ao": dwPYM(391).code = -20292
    dwPYM(392).py = "Ang": dwPYM(392).code = -20295
    dwPYM(393).py = "An": dwPYM(393).code = -20304
    dwPYM(394).py = "Ai": dwPYM(394).code = -20317
    dwPYM(395).py = "A": dwPYM(395).code = -20319
End Sub
