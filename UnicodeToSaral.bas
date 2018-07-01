'Licensed to the Apache Software Foundation (ASF) under one
'or more contributor license agreements.  See the NOTICE file
'distributed with this work for additional information
'regarding copyright ownership.  The ASF licenses this file
'to you under the Apache License, Version 2.0 (the
'"License"); you may not use this file except in compliance
'with the License.  You may obtain a copy of the License at
'
'  http://www.apache.org/licenses/LICENSE-2.0
'
'Unless required by applicable law or agreed to in writing,
'software distributed under the License is distributed on an
'"AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
'KIND, either express or implied.  See the License for the
'specific language governing permissions and limitations
'under the License.


Function mergeArrays(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
'joins 2 arrays

 Dim arr3() As Variant
 ReDim arr3(UBound(arr1) + UBound(arr2) + 1)

    Dim i As Integer
    Dim j As Integer
    For i = 0 To UBound(arr1)
        arr3(i) = arr1(i)
    Next i
    For j = 0 To UBound(arr2)
        arr3(UBound(arr1) + j) = arr2(j)
    Next j
 
 mergeArrays = arr3
End Function

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function
    


Sub UnicodeToSaral()
    'converts Unicode Characters to Saral
    'Use in conjuction with Saral to Unicode version 2 Excel
    
    Set RngTxt = Selection.Range
    Set RngFnd = RngTxt.Duplicate
    
    '›   ™   È   É   Ê   Ë   Ì   Í   Î   Ï   Ð   Ñ   Ò   Ó   Ô   Õ   Ö   ×   Ø   Ù   Ú   Û   Ü   Ý   Þ   ß   à   á   â   ã   ä   å   '   "   "   '
    level0_Saral = Array(ChrW(8250), ChrW(8482), ChrW(200), ChrW(201), ChrW(202), ChrW(203), ChrW(204), ChrW(205), ChrW(206), ChrW(207), ChrW(208), ChrW(209), ChrW(210), ChrW(211), ChrW(212), ChrW(213), ChrW(214), ChrW(215), ChrW(216), ChrW(217), ChrW(218), ChrW(219), ChrW(220), ChrW(221), ChrW(222), ChrW(223), ChrW(224), ChrW(225), ChrW(226), ChrW(227), ChrW(228), ChrW(229), ChrW(39), ChrW(34), ChrW(34), ChrW(39))
    level0_Uni = Array(ChrW(2404), ChrW(2365), ChrW(2406), ChrW(2407), ChrW(2408), ChrW(2409), ChrW(2410), ChrW(2411), ChrW(2412), ChrW(2413), ChrW(2414), ChrW(2415), ChrW(48), ChrW(49), ChrW(50), ChrW(51), ChrW(52), ChrW(53), ChrW(54), ChrW(55), ChrW(56), ChrW(57), ChrW(40), ChrW(41), ChrW(42), ChrW(43), ChrW(45), ChrW(739), ChrW(8763), ChrW(47), ChrW(36), ChrW(59), ChrW(8216), ChrW(8221), ChrW(8220), ChrW(8217))
    
    For i = LBound(level0_Uni) To UBound(level0_Uni)
        RngFnd = Replace(RngFnd, level0_Uni(i), level0_Saral(i))
    Next
    
    '+   =   )   9
    level1_Saral = Array(ChrW(43), ChrW(61), ChrW(41), ChrW(57))
    level1_Uni = Array(ChrW(2352) & ChrW(2370), ChrW(2352) & ChrW(2369), ChrW(2332) & ChrW(2381) & ChrW(2334), ChrW(2325) & ChrW(2381) & ChrW(2359))
    
    For i = LBound(level1_Uni) To UBound(level1_Uni)
        RngFnd = Replace(RngFnd, level1_Uni(i), level1_Saral(i))
    Next
    
    ' ¾   ½   ¼   "   ¹   ·      µ   ´   ²   ±   ¯   ®   ­      "   ª   ©   ¨   §   ¦   ¥   ¤   £   Œ   ˜   ¤   £   Œ   Š   -   "   "   ‡   …   ‰   ‹
    level2_Saral = Array(ChrW(190), ChrW(189), ChrW(188), ChrW(187), ChrW(185), ChrW(183), ChrW(182), ChrW(181), ChrW(180), ChrW(178), ChrW(177), ChrW(175), ChrW(174), ChrW(173), ChrW(172), ChrW(171), ChrW(170), ChrW(169), ChrW(168), ChrW(167), ChrW(166), ChrW(165), ChrW(164), ChrW(163), ChrW(338), ChrW(732), ChrW(164), ChrW(163), ChrW(338), ChrW(352), ChrW(8212), ChrW(8220), ChrW(8221), ChrW(8225), ChrW(8230), ChrW(8240), ChrW(8249))
    level2_Uni = Array(ChrW(2361) & ChrW(2381) & ChrW(2351), ChrW(2361) & ChrW(2381) & ChrW(2350), ChrW(2361) & ChrW(2381) & ChrW(2354), ChrW(2361) & ChrW(2381) & ChrW(2344), ChrW(2361) & ChrW(2371), ChrW(2347) & ChrW(2371), ChrW(2344) & ChrW(2381) & ChrW(2344), ChrW(2342) & ChrW(2381) & ChrW(2357), ChrW(2342) & ChrW(2381) & ChrW(2351), ChrW(2342) & ChrW(2381) & ChrW(2343), ChrW(2342) & ChrW(2381) & ChrW(2342), ChrW(2325) & ChrW(2381) & ChrW(2340), ChrW(2338) & ChrW(2381) & ChrW(2338), ChrW(2337) & ChrW(2381) & ChrW(2337), ChrW(2336) & ChrW(2381) & ChrW(2336), _
    ChrW(2335) & ChrW(2381) & ChrW(2351), ChrW(2335) & ChrW(2381) & ChrW(2336), ChrW(2335) & ChrW(2381) & ChrW(2335), ChrW(2335) & ChrW(2381) & ChrW(2352), ChrW(2329) & ChrW(2381) & ChrW(2350), ChrW(2329) & ChrW(2381) & ChrW(2328), ChrW(2329) & ChrW(2381) & ChrW(2327), ChrW(2329) & ChrW(2381) & ChrW(2326), ChrW(2329) & ChrW(2381) & ChrW(2325), ChrW(2309) & ChrW(2305), ChrW(2405), ChrW(2329) & ChrW(2381) & ChrW(2326), ChrW(2329) & ChrW(2381) & ChrW(2325), ChrW(2309) & ChrW(2305), ChrW(2309) & ChrW(2306), ChrW(2384), ChrW(2338) & ChrW(2364), _
    ChrW(2340) & ChrW(2381) & ChrW(2352) & ChrW(2381), ChrW(2340) & ChrW(2381) & ChrW(2352) & ChrW(2381) & ChrW(2346) & ChrW(2381) & ChrW(2352) & ChrW(2380), ChrW(2340) & ChrW(2381) & ChrW(2352) & ChrW(2381) & ChrW(2346) & ChrW(2381) & ChrW(2352) & ChrW(2379), ChrW(2340) & ChrW(2381) & ChrW(2352) & ChrW(2381) & ChrW(2346) & ChrW(2381) & ChrW(2352) & ChrW(2306), ChrW(2340) & ChrW(2381) & ChrW(2352) & ChrW(2381) & ChrW(2346) & ChrW(2381) & ChrW(2352) & ChrW(2305))
    
    'kaR qaR gaR 6aR ƒaR caR 7aR jaR zaR 'aR 3aR #aR DaR !aR `aR taR 4aR daR 2aR naR n_aR    paR faR baR waR maR yaR raR r_aR    laR vaR xaR 8aR saR haR ;aR ;_aR    9aR )aR
    level3_Saral = Array( _
    ChrW(107) & ChrW(97) & ChrW(82), ChrW(113) & ChrW(97) & ChrW(82), ChrW(103) & ChrW(97) & ChrW(82), ChrW(54) & ChrW(97) & ChrW(82), ChrW(402) & ChrW(97) & ChrW(82), ChrW(99) & ChrW(97) & ChrW(82), ChrW(55) & ChrW(97) & ChrW(82), ChrW(106) & ChrW(97) & ChrW(82), ChrW(122) & ChrW(97) & ChrW(82), ChrW(8216) & ChrW(97) & ChrW(82), ChrW(51) & ChrW(97) & ChrW(82), ChrW(35) & ChrW(97) & ChrW(82), ChrW(68) & ChrW(97) & ChrW(82), ChrW(33) & ChrW(97) & ChrW(82), ChrW(96) & ChrW(97) & ChrW(82), ChrW(116) & ChrW(97) & ChrW(82), ChrW(52) & ChrW(97) & ChrW(82), ChrW(100) & ChrW(97) & ChrW(82), ChrW(50) & ChrW(97) & ChrW(82), ChrW(110) & ChrW(97) & ChrW(82), ChrW(110) & ChrW(95) & ChrW(97) & ChrW(82), ChrW(112) & ChrW(97) & ChrW(82), ChrW(102) & ChrW(97) & ChrW(82), ChrW(98) & ChrW(97) & ChrW(82), ChrW(119) & ChrW(97) & ChrW(82), ChrW(109) & ChrW(97) & ChrW(82), ChrW(121) & ChrW(97) & ChrW(82), ChrW(114) & ChrW(97) & ChrW(82), _
    ChrW(114) & ChrW(95) & ChrW(97) & ChrW(82), ChrW(108) & ChrW(97) & ChrW(82), ChrW(118) & ChrW(97) & ChrW(82), ChrW(120) & ChrW(97) & ChrW(82), ChrW(56) & ChrW(97) & ChrW(82), ChrW(115) & ChrW(97) & ChrW(82), ChrW(104) & ChrW(97) & ChrW(82), ChrW(59) & ChrW(97) & ChrW(82), ChrW(59) & ChrW(95) & ChrW(97) & ChrW(82), ChrW(57) & ChrW(97) & ChrW(82), ChrW(41) & ChrW(97) & ChrW(82))
    level3_Uni = Array( _
    ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2326) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2327) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2328) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2329) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2330) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2331) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2333) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2334) & ChrW(2366), _
    ChrW(2352) & ChrW(2381) & ChrW(2335) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2336) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2337) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2338) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2339) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2340) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2341) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2342) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2343) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2344) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2345) & ChrW(2366), _
    ChrW(2352) & ChrW(2381) & ChrW(2346) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2347) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2348) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2349) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2350) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2351) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2352) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2353) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2354) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2357) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2358) & ChrW(2366), _
    ChrW(2352) & ChrW(2381) & ChrW(2359) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2360) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2361) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2355) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2356) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(2381) & ChrW(2359) & ChrW(2366), ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(2381) & ChrW(2334) & ChrW(2366))

    
    'kR  qR  gR  6R  ƒR  cR  7R  jR  zR  'R  3R  #R  DR  !R  `R  tR  4R  dR  2R  nR  n_R pR  fR  bR  wR  mR  yR  rR  r_R lR  vR  xR  8R  sR  hR  ;R  ;_R 9R  )R
    Dim level4_Saral() As Variant
    level4_Saral = Array(ChrW(107) & ChrW(82), ChrW(113) & ChrW(82), ChrW(103) & ChrW(82), ChrW(54) & ChrW(82), ChrW(402) & ChrW(82), ChrW(99) & ChrW(82), ChrW(55) & ChrW(82), ChrW(106) & ChrW(82), ChrW(122) & ChrW(82), ChrW(8216) & ChrW(82), ChrW(51) & ChrW(82), ChrW(35) & ChrW(82), ChrW(68) & ChrW(82), ChrW(33) & ChrW(82), ChrW(96) & ChrW(82), ChrW(116) & ChrW(82), ChrW(52) & ChrW(82), ChrW(100) & ChrW(82), ChrW(50) & ChrW(82), ChrW(110) & ChrW(82), ChrW(110) & ChrW(95) & ChrW(82), ChrW(112) & ChrW(82), ChrW(102) & ChrW(82), ChrW(98) & ChrW(82), ChrW(119) & ChrW(82), ChrW(109) & ChrW(82), ChrW(121) & ChrW(82), ChrW(114) & ChrW(82), ChrW(114) & ChrW(95) & ChrW(82), ChrW(108) & ChrW(82), ChrW(118) & ChrW(82), ChrW(120) & ChrW(82), ChrW(56) & ChrW(82), ChrW(115) & ChrW(82), ChrW(104) & ChrW(82), ChrW(59) & ChrW(82), ChrW(59) & ChrW(95) & ChrW(82), ChrW(57) & ChrW(82), ChrW(41) & ChrW(82))
    Dim level4_Uni() As Variant
    level4_Uni = Array(ChrW(2352) & ChrW(2381) & ChrW(2325), ChrW(2352) & ChrW(2381) & ChrW(2326), ChrW(2352) & ChrW(2381) & ChrW(2327), ChrW(2352) & ChrW(2381) & ChrW(2328), ChrW(2352) & ChrW(2381) & ChrW(2329), ChrW(2352) & ChrW(2381) & ChrW(2330), ChrW(2352) & ChrW(2381) & ChrW(2331), ChrW(2352) & ChrW(2381) & ChrW(2332), ChrW(2352) & ChrW(2381) & ChrW(2333), ChrW(2352) & ChrW(2381) & ChrW(2334), ChrW(2352) & ChrW(2381) & ChrW(2335), ChrW(2352) & ChrW(2381) & ChrW(2336), ChrW(2352) & ChrW(2381) & ChrW(2337), ChrW(2352) & ChrW(2381) & ChrW(2338), _
    ChrW(2352) & ChrW(2381) & ChrW(2339), ChrW(2352) & ChrW(2381) & ChrW(2340), ChrW(2352) & ChrW(2381) & ChrW(2341), ChrW(2352) & ChrW(2381) & ChrW(2342), ChrW(2352) & ChrW(2381) & ChrW(2343), ChrW(2352) & ChrW(2381) & ChrW(2344), ChrW(2352) & ChrW(2381) & ChrW(2345), ChrW(2352) & ChrW(2381) & ChrW(2346), ChrW(2352) & ChrW(2381) & ChrW(2347), ChrW(2352) & ChrW(2381) & ChrW(2348), ChrW(2352) & ChrW(2381) & ChrW(2349), ChrW(2352) & ChrW(2381) & ChrW(2350), ChrW(2352) & ChrW(2381) & ChrW(2351), ChrW(2352) & ChrW(2381) & ChrW(2352), ChrW(2352) & ChrW(2381) & ChrW(2353), _
    ChrW(2352) & ChrW(2381) & ChrW(2354), ChrW(2352) & ChrW(2381) & ChrW(2357), ChrW(2352) & ChrW(2381) & ChrW(2358), ChrW(2352) & ChrW(2381) & ChrW(2359), ChrW(2352) & ChrW(2381) & ChrW(2360), ChrW(2352) & ChrW(2381) & ChrW(2361), ChrW(2352) & ChrW(2381) & ChrW(2355), ChrW(2352) & ChrW(2381) & ChrW(2356), ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(2381) & ChrW(2359), ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(2381) & ChrW(2334))

    'kIR.    qIR.    gIR.    6IR.    ƒIR.    cIR.    7IR.    jIR.    zIR.    'IR.    3IR.    #IR.    DIR.    !IR.    `IR.    tIR.    4IR.    dIR.    2IR.    nIR.    n_IR.   pIR.    fIR.    bIR.    wIR.    mIR.    yIR.    rIR.    r_IR.   lIR.    vIR.    xIR.    8IR.    sIR.    hIR.    ;IR.    ;_IR.   9IR.    )IR.
    Dim level5_Saral() As Variant
    level5_Saral = Array( _
    ChrW(107) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(113) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(103) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(54) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(402) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(99) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(55) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(106) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(122) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(8216) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(51) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(35) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(68) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(33) & ChrW(73) & ChrW(82) & ChrW(46), _
ChrW(96) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(116) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(52) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(100) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(50) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(110) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(110) & ChrW(95) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(112) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(102) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(98) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(119) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(109) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(121) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(114) & ChrW(73) & ChrW(82) & ChrW(46), _
ChrW(114) & ChrW(95) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(108) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(118) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(120) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(56) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(115) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(104) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(59) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(59) & ChrW(95) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(57) & ChrW(73) & ChrW(82) & ChrW(46), ChrW(41) & ChrW(73) & ChrW(82) & ChrW(46))
    Dim level5_Uni() As Variant
     level5_Uni = Array( _
    ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2326) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2327) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2328) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2329) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2330) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2331) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2333) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2334) & ChrW(2368) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2335) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2336) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2337) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2338) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2339) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2340) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2341) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2342) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2343) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2344) & ChrW(2368) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2345) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2346) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2347) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2348) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2349) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2350) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2351) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2352) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2353) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2354) & ChrW(2368) & ChrW(2306), _
    ChrW(2352) & ChrW(2381) & ChrW(2357) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2358) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2359) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2360) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2361) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2355) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2356) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2325) & ChrW(2381) & ChrW(2359) & ChrW(2368) & ChrW(2306), ChrW(2352) & ChrW(2381) & ChrW(2332) & ChrW(2381) & ChrW(2334) & ChrW(2368) & ChrW(2306))


    '¢   q/  g/  6/  ƒ/  c/  7/  j/  z/  '/  3/  #/  ¿   !/  `/  5   4/  ³   2/  n/  n_/ p/  ¸   b/  w/  m/  y/  rR  r_R l/  v/  &   8/  s/  º   ;/  ;_/ 9/  )/
    Dim level6_Saral() As Variant
    level6_Saral = Array(ChrW(162), ChrW(113) & ChrW(47), ChrW(103) & ChrW(47), ChrW(54) & ChrW(47), ChrW(402) & ChrW(47), ChrW(99) & ChrW(47), ChrW(55) & ChrW(47), ChrW(106) & ChrW(47), ChrW(122) & ChrW(47), ChrW(8216) & ChrW(47), ChrW(51) & ChrW(47), ChrW(35) & ChrW(47), ChrW(191), ChrW(33) & ChrW(47), ChrW(96) & ChrW(47), ChrW(53), ChrW(52) & ChrW(47), ChrW(179), ChrW(50) & ChrW(47), ChrW(110) & ChrW(47), _
    ChrW(110) & ChrW(95) & ChrW(47), ChrW(112) & ChrW(47), ChrW(184), ChrW(98) & ChrW(47), ChrW(119) & ChrW(47), ChrW(109) & ChrW(47), ChrW(121) & ChrW(47), ChrW(114) & ChrW(82), ChrW(114) & ChrW(95) & ChrW(82), ChrW(108) & ChrW(47), ChrW(118) & ChrW(47), ChrW(38), ChrW(56) & ChrW(47), ChrW(115) & ChrW(47), ChrW(186), ChrW(59) & ChrW(47), ChrW(59) & ChrW(95) & ChrW(47), ChrW(57) & ChrW(47), ChrW(41) & ChrW(47))
    Dim level6_Uni() As Variant
    level6_Uni = Array(ChrW(2325) & ChrW(2381) & ChrW(2352), ChrW(2326) & ChrW(2381) & ChrW(2352), ChrW(2327) & ChrW(2381) & ChrW(2352), ChrW(2328) & ChrW(2381) & ChrW(2352), ChrW(2329) & ChrW(2381) & ChrW(2352), ChrW(2330) & ChrW(2381) & ChrW(2352), ChrW(2331) & ChrW(2381) & ChrW(2352), ChrW(2332) & ChrW(2381) & ChrW(2352), ChrW(2333) & ChrW(2381) & ChrW(2352), ChrW(2334) & ChrW(2381) & ChrW(2352), ChrW(2335) & ChrW(2381) & ChrW(2352), ChrW(2336) & ChrW(2381) & ChrW(2352), ChrW(2337) & ChrW(2381) & ChrW(2352), ChrW(2338) & ChrW(2381) & ChrW(2352), ChrW(2339) & ChrW(2381) & ChrW(2352), ChrW(2340) & ChrW(2381) & ChrW(2352), ChrW(2341) & ChrW(2381) & ChrW(2352), _
    ChrW(2342) & ChrW(2381) & ChrW(2352), ChrW(2343) & ChrW(2381) & ChrW(2352), ChrW(2344) & ChrW(2381) & ChrW(2352), ChrW(2345) & ChrW(2381) & ChrW(2352), ChrW(2346) & ChrW(2381) & ChrW(2352), ChrW(2347) & ChrW(2381) & ChrW(2352), ChrW(2348) & ChrW(2381) & ChrW(2352), ChrW(2349) & ChrW(2381) & ChrW(2352), ChrW(2350) & ChrW(2381) & ChrW(2352), ChrW(2351) & ChrW(2381) & ChrW(2352), ChrW(2352) & ChrW(2381) & ChrW(2352), ChrW(2353) & ChrW(2381) & ChrW(2352), ChrW(2354) & ChrW(2381) & ChrW(2352), ChrW(2357) & ChrW(2381) & ChrW(2352), ChrW(2358) & ChrW(2381) & ChrW(2352), ChrW(2359) & ChrW(2381) & ChrW(2352), ChrW(2360) & ChrW(2381) & ChrW(2352), ChrW(2361) & ChrW(2381) & ChrW(2352), ChrW(2355) & ChrW(2381) & ChrW(2352), ChrW(2356) & ChrW(2381) & ChrW(2352), ChrW(2325) & ChrW(2381) & ChrW(2359) & ChrW(2381) & ChrW(2352), ChrW(2332) & ChrW(2381) & ChrW(2334) & ChrW(2381) & ChrW(2352))

    
  'K   Q   G   ^   ƒ\  C   7\  J   Z   '   3\  #\  D\  !\  ~   T   $   !\  @   N   N_  P   F   B   W   M   Y   -   -_  L   V   X   *   S   H   ;\  ;_\ (   o
    Dim level7_Saral() As Variant
    level7_Saral = Array(ChrW(75), ChrW(81), ChrW(71), ChrW(94), ChrW(402) & ChrW(92), ChrW(67), ChrW(55) & ChrW(92), ChrW(74), ChrW(90), ChrW(8217), ChrW(51) & ChrW(92), ChrW(35) & ChrW(92), ChrW(68) & ChrW(92), ChrW(33) & ChrW(92), ChrW(126), ChrW(84), ChrW(36), ChrW(33) & ChrW(92), ChrW(64), ChrW(78), ChrW(78) & ChrW(95), ChrW(80), ChrW(70), ChrW(66), ChrW(87), ChrW(77), ChrW(89), ChrW(8211), ChrW(8211) & ChrW(95), ChrW(76), ChrW(86), ChrW(88), ChrW(42), ChrW(83), ChrW(72), ChrW(59) & ChrW(92), ChrW(59) & ChrW(95) & ChrW(92), ChrW(40), ChrW(8226))
    Dim level7_Uni() As Variant
    level7_Uni = Array( _
    ChrW(2325) & ChrW(2381), ChrW(2326) & ChrW(2381), ChrW(2327) & ChrW(2381), ChrW(2328) & ChrW(2381), ChrW(2329) & ChrW(2381), ChrW(2330) & ChrW(2381), ChrW(2331) & ChrW(2381), ChrW(2332) & ChrW(2381), ChrW(2333) & ChrW(2381), ChrW(2334) & ChrW(2381), ChrW(2335) & ChrW(2381), ChrW(2336) & ChrW(2381), ChrW(2337) & ChrW(2381), ChrW(2338) & ChrW(2381), ChrW(2339) & ChrW(2381), ChrW(2340) & ChrW(2381), ChrW(2341) & ChrW(2381), ChrW(2342) & ChrW(2381), ChrW(2343) & ChrW(2381), ChrW(2344) & ChrW(2381), ChrW(2345) & ChrW(2381), ChrW(2346) & ChrW(2381), ChrW(2347) & ChrW(2381), ChrW(2348) & ChrW(2381), ChrW(2349) & ChrW(2381), ChrW(2350) & ChrW(2381), ChrW(2351) & ChrW(2381), _
    ChrW(2352) & ChrW(2381), ChrW(2353) & ChrW(2381), ChrW(2354) & ChrW(2381), ChrW(2357) & ChrW(2381), ChrW(2358) & ChrW(2381), ChrW(2359) & ChrW(2381), ChrW(2360) & ChrW(2381), ChrW(2361) & ChrW(2381), ChrW(2355) & ChrW(2381), ChrW(2356) & ChrW(2381), ChrW(2325) & ChrW(2381) & ChrW(2359) & ChrW(2381), ChrW(2332) & ChrW(2381) & ChrW(2334) & ChrW(2381))
    
  'k   q   g   6   ƒ   c   7   j   z   '   3   #   D   !   `   t   4   d   2   n   n_  p   f   b   w   m   y   r   r_  l   v   x   8   s   h   ;   ;_  9   )
    Dim level8_Saral() As Variant
    level8_Saral = Array(ChrW(107), ChrW(113), ChrW(103), ChrW(54), ChrW(402), ChrW(99), ChrW(55), ChrW(106), ChrW(122), ChrW(8216), ChrW(51), ChrW(35), ChrW(68), ChrW(33), ChrW(96), ChrW(116), ChrW(52), ChrW(100), ChrW(50), ChrW(110), ChrW(110) & ChrW(95), ChrW(112), ChrW(102), ChrW(98), ChrW(119), ChrW(109), ChrW(121), ChrW(114), ChrW(114) & ChrW(95), ChrW(108), ChrW(118), ChrW(120), ChrW(56), ChrW(115), ChrW(104), ChrW(59), ChrW(59) & ChrW(95), ChrW(57), ChrW(41))
    Dim level8_Uni() As Variant
    level8_Uni = Array(ChrW(2325), ChrW(2326), ChrW(2327), ChrW(2328), ChrW(2329), ChrW(2330), ChrW(2331), ChrW(2332), ChrW(2333), ChrW(2334), ChrW(2335), ChrW(2336), ChrW(2337), ChrW(2338), ChrW(2339), ChrW(2340), ChrW(2341), ChrW(2342), ChrW(2343), ChrW(2344), ChrW(2345), ChrW(2346), ChrW(2347), ChrW(2348), ChrW(2349), ChrW(2350), ChrW(2351), ChrW(2352), ChrW(2353), ChrW(2354), ChrW(2357), ChrW(2358), ChrW(2359), ChrW(2360), ChrW(2361), ChrW(2355), ChrW(2356), ChrW(2325) & ChrW(2381) & ChrW(2359), ChrW(2332) & ChrW(2381) & ChrW(2334))
    
    'A   Aa  [   {   ]   }   0   0e  1   1<  l<  l<  0   0e  †   ˆ   Š   A:
    Dim level9_Saral() As Variant
    level9_Saral = Array(ChrW(65), ChrW(65) & ChrW(97), ChrW(91), ChrW(123), ChrW(93), ChrW(125), ChrW(48), ChrW(48) & ChrW(101), ChrW(49), ChrW(49) & ChrW(60), ChrW(108) & ChrW(60), ChrW(108) & ChrW(60), ChrW(48), ChrW(48) & ChrW(101), ChrW(8224), ChrW(710), ChrW(352), ChrW(65) & ChrW(58), ChrW(65) & ChrW(97) & ChrW(353))
    Dim level9_Uni() As Variant
    level9_Uni = Array(ChrW(2309), ChrW(2310), ChrW(2311), ChrW(2312), ChrW(2313), ChrW(2314), ChrW(2319), ChrW(2320), ChrW(2315), ChrW(2315) & ChrW(2371), ChrW(2316), ChrW(2401), ChrW(2319), ChrW(2320), ChrW(2323), ChrW(2324), ChrW(2309) & ChrW(2306), ChrW(2309) & ChrW(2307), ChrW(2321))
    
    '_   -   |   \   E   e   u   U   I   i   O   o   a   :   <   ,   >   .   ›   .
    Dim level10_Saral() As Variant
    level10_Saral = Array(ChrW(8250), ChrW(95), ChrW(45), ChrW(124), ChrW(92), ChrW(69), ChrW(101), ChrW(117), ChrW(85), ChrW(73), ChrW(105), ChrW(79), ChrW(111), ChrW(97), ChrW(58), ChrW(60), ChrW(44), ChrW(62), ChrW(46), ChrW(8250), ChrW(46), ChrW(353), ChrW(97) & ChrW(353))
    Dim level10_Uni() As Variant
    level10_Uni = Array(ChrW(124), ChrW(2364), ChrW(45), ChrW(46), ChrW(2381), ChrW(2376), ChrW(2375), ChrW(2369), ChrW(2370), ChrW(2368), ChrW(2367), ChrW(2380), ChrW(2379), ChrW(2366), ChrW(2307), ChrW(2371), ChrW(44), ChrW(2305), ChrW(2417), ChrW(2404), ChrW(2306), ChrW(2373), ChrW(2377))
    
    
    
    fndList = mergeArrays(level1_Uni, level2_Uni)
    fndList = mergeArrays(fndList, level3_Uni)
    fndList = mergeArrays(fndList, level4_Uni)
    fndList = mergeArrays(fndList, level5_Uni)
    fndList = mergeArrays(fndList, level6_Uni)
    fndList = mergeArrays(fndList, level7_Uni)
    fndList = mergeArrays(fndList, level8_Uni)
    fndList = mergeArrays(fndList, level9_Uni)
    fndList = mergeArrays(fndList, level10_Uni)

    rplcList = mergeArrays(level1_Saral, level2_Saral)
    rplcList = mergeArrays(rplcList, level3_Saral)
    rplcList = mergeArrays(rplcList, level4_Saral)
    rplcList = mergeArrays(rplcList, level5_Saral)
    rplcList = mergeArrays(rplcList, level6_Saral)
    rplcList = mergeArrays(rplcList, level7_Saral)
    rplcList = mergeArrays(rplcList, level8_Saral)
    rplcList = mergeArrays(rplcList, level9_Saral)
    rplcList = mergeArrays(rplcList, level10_Saral)
    '094D position change
    Dim buff() As String
    ReDim buff(Len(RngFnd) - 1)
    For i = 1 To Len(RngFnd)
        buff(i - 1) = Mid$(RngFnd, i, 1)
    Next
    
    '2381 velanti replace
    For i = LBound(buff) To UBound(buff)
    
    If (AscW(buff(i)) = 2367) Then
    If (i - 2 > 0) Then
        If (AscW(buff(i - 2)) = 2381) Then
        'jodakshar aahe
            If (i - 4 > 0 & AscW(buff(i - 4)) = 2381) Then
            'double jodakshar
            'akshar | 094D | akshar | 094D | akshar | velanti
            Dim zemp1 As String
            Dim zemp2 As String
            Dim zemp3 As String
            Dim zemp4 As String
            Dim zemp5 As String
            Dim zemp6 As String
            zemp6 = buff(i)
            zemp5 = buff(i - 1)
            zemp4 = buff(i - 2)
            zemp3 = buff(i - 3)
            zemp2 = buff(i - 4)
            zemp1 = buff(i - 5)
            buff(i) = zemp5
            buff(i - 1) = zemp4
            buff(i - 2) = zemp3
            buff(i - 3) = zemp2
            buff(i - 4) = zemp1
            buff(i - 5) = zemp6
            
            Else
            'single dokakshar
            'akshar | 094D | akshar | velanti
            Dim temp1 As String
            Dim temp2 As String
            Dim temp3 As String
            Dim temp4 As String
            temp4 = buff(i)
            temp3 = buff(i - 1)
            temp2 = buff(i - 2)
            temp1 = buff(i - 3)
            buff(i) = temp3
            buff(i - 1) = temp2
            buff(i - 2) = temp1
            buff(i - 3) = temp4
            End If
        Else
        'jodakshar naahi
        Dim temp5 As String
        temp5 = buff(i - 1)
        buff(i - 1) = buff(i)
        buff(i) = temp5
        End If
    Else
    Dim temp As String 'Pahilya aksharala velanti
    temp = buff(i - 1)
    buff(i - 1) = buff(i)
    buff(i) = temp
    End If
    End If
    Next
    
    RngFnd = Join(buff, "")
    

    For i = LBound(fndList) To UBound(fndList)
        RngFnd = Replace(RngFnd, fndList(i), rplcList(i))
    Next


    'Restore Numbers & Symbols
    level0Correct = Array(ChrW(8250), ChrW(8482), ChrW(200), ChrW(201), ChrW(202), ChrW(203), ChrW(204), ChrW(205), ChrW(206), ChrW(207), ChrW(208), ChrW(209), ChrW(210), ChrW(211), ChrW(212), ChrW(213), ChrW(214), ChrW(215), ChrW(216), ChrW(217), ChrW(218), ChrW(219), ChrW(220), ChrW(221), ChrW(222), ChrW(223), ChrW(224), ChrW(225), ChrW(226), ChrW(227), ChrW(228), ChrW(229), ChrW(39), ChrW(34), ChrW(34), ChrW(39))
    
    For i = LBound(level0_Saral) To UBound(level0_Saral)
        RngFnd = Replace(RngFnd, level0_Saral(i), level0Correct(i))
    Next


    RngTxt.Text = RngFnd
    Selection.ClearFormatting
    Selection.Font.Name = "Marathi Saral-3"

    
End Sub


