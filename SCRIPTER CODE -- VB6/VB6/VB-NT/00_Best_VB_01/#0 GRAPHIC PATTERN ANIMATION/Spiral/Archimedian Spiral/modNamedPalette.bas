Attribute VB_Name = "modNamedPalette"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''' AUTHOR:       Nathan Moschkin                               '''''''
''''''' MODULE:       modCDNamedPalette                             '''''''
''''''' DESCRIPTION:  Create and manage the named palette           '''''''
'''''''                                                             '''''''
''''''' COPYRIGHT:    Datavations, LLC  2003                        '''''''
'''''''                                                             '''''''
'''''''                                                             '''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Option Base 0

Public Type COLORINFO
    ColorName As String
    Value As Long
End Type

Public g_Colors() As COLORINFO
Public g_ColorCount As Long

Private Sub AddColorItem(ByVal ColorName As String, ByVal Value As Long)

    ReDim Preserve g_Colors(g_ColorCount)
    
    With g_Colors(g_ColorCount)
        .ColorName = ColorName
        .Value = Value
    
    End With
    
    g_ColorCount = g_ColorCount + 1
    
End Sub

Public Sub ListNamedPalette()

    Erase g_Colors
    g_ColorCount = 0
    
    AddColorItem "Black", 0
    AddColorItem "Alice Blue", 16775408
    AddColorItem "Blue Violet", 14822282
    AddColorItem "Cadet Blue", 10526303
    AddColorItem "Cadet Blue 1", 16774552
    AddColorItem "Cadet Blue 2", 15656334
    AddColorItem "Cadet Blue 3", 13485434
    AddColorItem "Cadet Blue 4", 9143891
    AddColorItem "Cornflower Blue", 15570276
    AddColorItem "Darkslate Blue", 9125192
    AddColorItem "Dark Turquoise", 13749760
    AddColorItem "Deep Sky Blue", 16760576
    AddColorItem "Deep Sky Blue 2", 15643136
    AddColorItem "Deep Sky Blue 3", 13474304
    AddColorItem "Deep Sky Blue 4", 9136128
    AddColorItem "Dodger Blue", 16748574
    AddColorItem "Dodger Blue 2", 15631900
    AddColorItem "Dodger Blue 3", 13464600
    AddColorItem "Dodger Blue 4", 9129488
    AddColorItem "Light Blue", 15128749
    AddColorItem "Light Blue 1", 16773055
    AddColorItem "Light Blue 2", 15654834
    AddColorItem "Light Blue 3", 13484186
    AddColorItem "Light Blue 4", 9143144
    AddColorItem "Light Cyan", 16777184
    AddColorItem "Light Cyan 2", 15658705
    AddColorItem "Light Cyan 3", 13487540
    AddColorItem "Light Cyan 4", 9145210
    AddColorItem "Light Sky Blue", 16436871
    AddColorItem "Light Sky Blue 1", 16769712
    AddColorItem "Light Sky Blue 2", 15651748
    AddColorItem "Light Sky Blue 3", 13481613
    AddColorItem "Light Sky Blue 4", 9141088
    AddColorItem "Light Slate Blue", 16740484
    AddColorItem "Light Steel Blue", 14599344
    AddColorItem "Light Steel Blue 1", 16769482
    AddColorItem "Light Steel Blue 2", 15651516
    AddColorItem "Light Steel Blue 3", 13481378
    AddColorItem "Light Steel Blue 4", 9141102
    AddColorItem "Medium Aquamarine", 11193702
    AddColorItem "Medium Blue", 13434880
    AddColorItem "Medium Slate Blue", 15624315
    AddColorItem "Medium Turquoise", 13422920
    AddColorItem "Midnight Blue", 7346457
    AddColorItem "Navy Blue", 8388608
    AddColorItem "Pale Turquoise", 15658671
    AddColorItem "Pale Turquoise 1", 16777147
    AddColorItem "Pale Turquoise 2", 15658670
    AddColorItem "Pale Turquoise 3", 13487510
    AddColorItem "Pale Turquoise 4", 9145190
    AddColorItem "Powder Blue", 15130800
    AddColorItem "Royal Blue", 14772545
    AddColorItem "Royal Blue 1", 16741960
    AddColorItem "Royal Blue 2", 15625795
    AddColorItem "Royal Blue 3", 13459258
    AddColorItem "Royal Blue 4", 9125927
    AddColorItem "Sky Blue", 15453831
    AddColorItem "Sky Blue 1", 16764551
    AddColorItem "Sky Blue 2", 15646846
    AddColorItem "Sky Blue 3", 13477484
    AddColorItem "Sky Blue 4", 9138250
    AddColorItem "Slate Blue", 13458026
    AddColorItem "Slate Blue 1", 16740227
    AddColorItem "Slate Blue 2", 15624058
    AddColorItem "Slate Blue 3", 13457769
    AddColorItem "Slate Blue 4", 9124935
    AddColorItem "Steel Blue", 11829830
    AddColorItem "Steel Blue 1", 16758883
    AddColorItem "Steel Blue 2", 15641692
    AddColorItem "Steel Blue 3", 13472847
    AddColorItem "Steel Blue 4", 9135158
    AddColorItem "Aquamarine", 13959039
    AddColorItem "Aquamarine 2", 13037174
    AddColorItem "Aquamarine 4", 7637829
    AddColorItem "Azure", 16777200
    AddColorItem "Azure 2", 15658720
    AddColorItem "Azure 3", 13487553
    AddColorItem "Azure 4", 9145219
    AddColorItem "Blue", 16711680
    AddColorItem "Blue 2", 15597568
    AddColorItem "Blue 4", 9109504
    AddColorItem "Cyan", 16776960
    AddColorItem "Cyan 2", 15658496
    AddColorItem "Cyan 3", 13487360
    AddColorItem "Cyan 4", 9145088
    AddColorItem "Turquoise", 13688896
    AddColorItem "Turquoise 1", 16774400
    AddColorItem "Turquoise 2", 15656192
    AddColorItem "Turquoise 3", 13485312
    AddColorItem "Rosy Brown", 9408444
    AddColorItem "Rosy Brown 1", 12698111
    AddColorItem "Rosy Brown 2", 11842798
    AddColorItem "Rosy Brown 3", 10197965
    AddColorItem "Rosy Brown 4", 6908299
    AddColorItem "Saddle Brown", 1262987
    AddColorItem "Sandy Brown", 6333684
    AddColorItem "Beige", 14480885
    AddColorItem "Brown", 2763429
    AddColorItem "Brown 1", 4210943
    AddColorItem "Brown 2", 3881966
    AddColorItem "Brown 3", 3355597
    AddColorItem "Brown 4", 2302859
    AddColorItem "Burlywood", 8894686
    AddColorItem "Burlywood 1", 10212351
    AddColorItem "Burlywood 2", 9553390
    AddColorItem "Burlywood 3", 8235725
    AddColorItem "Burlywood 4", 5600139
    AddColorItem "Chocolate", 1993170
    AddColorItem "Chocolate 1", 2392063
    AddColorItem "Chocolate 2", 2193134
    AddColorItem "Chocolate 3", 1926861
    AddColorItem "Peru", 4163021
    AddColorItem "Tan", 9221330
    AddColorItem "Tan 1", 5219839
    AddColorItem "Tan 2", 4823790
    AddColorItem "Dark Slate Gray", 5197615
    AddColorItem "Dark Slate Gray 1", 16777111
    AddColorItem "Dark Slate Gray 2", 15658637
    AddColorItem "Dark Slate Gray 3", 13487481
    AddColorItem "Dark Slate Gray 4", 9145170
    AddColorItem "Dim Gray", 6908265
    AddColorItem "Light Gray", 13882323
    AddColorItem "Light Slate Gray", 10061943
    AddColorItem "Slate Gray", 9470064
    AddColorItem "Slate Gray 1", 16769734
    AddColorItem "Slate Gray 2", 15651769
    AddColorItem "Slate Gray 3", 13481631
    AddColorItem "Slate Gray 4", 9141100
    AddColorItem "Grey", 12500670
    AddColorItem "Grey 1", 197379
    AddColorItem "Grey 2", 328965
    AddColorItem "Grey 3", 526344
    AddColorItem "Grey 4", 657930
    AddColorItem "Grey 5", 855309
    AddColorItem "Grey 6", 986895
    AddColorItem "Grey 7", 1184274
    AddColorItem "Grey 8", 1315860
    AddColorItem "Grey 9", 1513239
    AddColorItem "Grey 10", 1710618
    AddColorItem "Grey 11", 1842204
    AddColorItem "Grey 12", 2039583
    AddColorItem "Grey 13", 2171169
    AddColorItem "Grey 14", 2368548
    AddColorItem "Grey 15", 2500134
    AddColorItem "Grey 16", 2697513
    AddColorItem "Grey 17", 2829099
    AddColorItem "Grey 18", 3026478
    AddColorItem "Grey 19", 3158064
    AddColorItem "Grey 20", 3355443
    AddColorItem "Grey 21", 3552822
    AddColorItem "Grey 22", 3684408
    AddColorItem "Grey 23", 3881787
    AddColorItem "Grey 24", 4013373
    AddColorItem "Grey 25", 4210752
    AddColorItem "Grey 26", 4342338
    AddColorItem "Grey 27", 4539717
    AddColorItem "Grey 28", 4671303
    AddColorItem "Grey 29", 4868682
    AddColorItem "Grey 30", 5066061
    AddColorItem "Grey 31", 5197647
    AddColorItem "Grey 32", 5395026
    AddColorItem "Grey 33", 5526612
    AddColorItem "Grey 34", 5723991
    AddColorItem "Grey 35", 5855577
    AddColorItem "Grey 36", 6052956
    AddColorItem "Grey 37", 6184542
    AddColorItem "Grey 38", 6381921
    AddColorItem "Grey 39", 6513507
    AddColorItem "Grey 40", 6710886
    AddColorItem "Grey 42", 7039851
    AddColorItem "Grey 43", 7237230
    AddColorItem "Grey 44", 7368816
    AddColorItem "Grey 45", 7566195
    AddColorItem "Grey 46", 7697781
    AddColorItem "Grey 47", 7895160
    AddColorItem "Grey 48", 8026746
    AddColorItem "Grey 49", 8224125
    AddColorItem "Grey 50", 8355711
    AddColorItem "Grey 51", 8553090
    AddColorItem "Grey 52", 8750469
    AddColorItem "Grey 53", 8882055
    AddColorItem "Grey 54", 9079434
    AddColorItem "Grey 55", 9211020
    AddColorItem "Grey 56", 9408399
    AddColorItem "Grey 57", 9539985
    AddColorItem "Grey 58", 9737364
    AddColorItem "Grey 59", 9868950
    AddColorItem "Grey 60", 10066329
    AddColorItem "Grey 61", 10263708
    AddColorItem "Grey 62", 10395294
    AddColorItem "Grey 63", 10592673
    AddColorItem "Grey 64", 10724259
    AddColorItem "Grey 65", 10921638
    AddColorItem "Grey 66", 11053224
    AddColorItem "Grey 67", 11250603
    AddColorItem "Grey 68", 11382189
    AddColorItem "Grey 69", 11579568
    AddColorItem "Grey 70", 11776947
    AddColorItem "Grey 71", 11908533
    AddColorItem "Grey 72", 12105912
    AddColorItem "Grey 73", 12237498
    AddColorItem "Grey 74", 12434877
    AddColorItem "Grey 75", 12566463
    AddColorItem "Grey 76", 12763842
    AddColorItem "Grey 77", 12895428
    AddColorItem "Grey 78", 13092807
    AddColorItem "Grey 79", 13224393
    AddColorItem "Grey 80", 13421772
    AddColorItem "Grey 81", 13619151
    AddColorItem "Grey 82", 13750737
    AddColorItem "Grey 83", 13948116
    AddColorItem "Grey 84", 14079702
    AddColorItem "Grey 85", 14277081
    AddColorItem "Grey 86", 14408667
    AddColorItem "Grey 87", 14606046
    AddColorItem "Grey 88", 14737632
    AddColorItem "Grey 89", 14935011
    AddColorItem "Grey 90", 15066597
    AddColorItem "Grey 91", 15263976
    AddColorItem "Grey 92", 15461355
    AddColorItem "Grey 93", 15592941
    AddColorItem "Grey 94", 15790320
    AddColorItem "Grey 95", 15921906
    AddColorItem "Grey 96", 16119285
    AddColorItem "Grey 97", 16250871
    AddColorItem "Grey 98", 16448250
    AddColorItem "Grey 99", 16579836
    AddColorItem "Grey 100", 16777215
    AddColorItem "Dark Green", 25600
    AddColorItem "Dark Khaki", 7059389
    AddColorItem "Dark Olivegreen", 3107669
    AddColorItem "Dark Olivegreen 1", 7405514
    AddColorItem "Dark Olivegreen 2", 6876860
    AddColorItem "Dark Olivegreen 3", 5950882
    AddColorItem "Dark Olivegreen 4", 4033390
    AddColorItem "Dark Sea Green", 9419919
    AddColorItem "Dark Sea Green 1", 12713921
    AddColorItem "Dark Sea Green 2", 11857588
    AddColorItem "Dark Sea Green 3", 10210715
    AddColorItem "Dark Sea Green 4", 6916969
    AddColorItem "Forest Green", 2263842
    AddColorItem "Green Yellow", 3145645
    AddColorItem "Lawn Green", 64636
    AddColorItem "Light Sea Green", 11186720
    AddColorItem "Lime Green", 3329330
    AddColorItem "Medium Sea Green", 7451452
    AddColorItem "Medium Spring Green", 10156544
    AddColorItem "Mint Cream", 16449525
    AddColorItem "Olive Drab", 2330219
    AddColorItem "Olive Drab 1", 4128704
    AddColorItem "Olive Drab 2", 3862195
    AddColorItem "Olive Drab 3", 3329434
    AddColorItem "Olive Drab 4", 2263913
    AddColorItem "Pale Green", 10025880
    AddColorItem "Pale Green 1", 10157978
    AddColorItem "Pale Green 2", 9498256
    AddColorItem "Pale Green 3", 8179068
    AddColorItem "Pale Green 4", 5540692
    AddColorItem "Sea Green", 5737262
    AddColorItem "Sea Green 1", 10485588
    AddColorItem "Sea Green 2", 9760334
    AddColorItem "Sea Green 3", 8441155
    AddColorItem "Spring Green", 8388352
    AddColorItem "Spring Green 2", 7794176
    AddColorItem "Spring Green 3", 6737152
    AddColorItem "Spring Green 4", 4557568
    AddColorItem "Chartreuse", 65407
    AddColorItem "Chartreuse 2", 61046
    AddColorItem "Chartreuse 3", 52582
    AddColorItem "Chartreuse 4", 35653
    AddColorItem "Green", 65280
    AddColorItem "Green 2", 60928
    AddColorItem "Green 3", 52480
    AddColorItem "Green 4", 35584
    AddColorItem "Khaki", 9234160
    AddColorItem "Khaki 1", 9434879
    AddColorItem "Khaki 2", 8775406
    AddColorItem "Khaki 3", 7587533
    AddColorItem "Khaki 4", 5146251
    AddColorItem "Dark Orange", 36095
    AddColorItem "Dark Orange 1", 32767
    AddColorItem "Dark Orange 2", 30446
    AddColorItem "Dark Orange 3", 26317
    AddColorItem "Dark Orange 4", 17803
    AddColorItem "Dark Salmon", 8034025
    AddColorItem "Light Coral", 8421616
    AddColorItem "Light Salmon", 8036607
    AddColorItem "Light Salmon 2", 7509486
    AddColorItem "Light Salmon 3", 6455757
    AddColorItem "Light Salmon 4", 4347787
    AddColorItem "Peach Puff", 12180223
    AddColorItem "Peach Puff 2", 11389934
    AddColorItem "Peach Puff 3", 9809869
    AddColorItem "Peach Puff 4", 6649739
    AddColorItem "Bisque", 12903679
    AddColorItem "Bisque 2", 12047854
    AddColorItem "Bisque 3", 10401741
    AddColorItem "Bisque 4", 7044491
    AddColorItem "Coral", 5275647
    AddColorItem "Coral 1", 5665535
    AddColorItem "Coral 2", 5270254
    AddColorItem "Coral 3", 4545485
    AddColorItem "Coral 4", 3096203
    AddColorItem "Honeydew", 15794160
    AddColorItem "Honeydew 2", 14741216
    AddColorItem "Honeydew 3", 12701121
    AddColorItem "Honeydew 4", 8620931
    AddColorItem "Orange", 42495
    AddColorItem "Orange 2", 39662
    AddColorItem "Orange 3", 34253
    AddColorItem "Orange 4", 23179
    AddColorItem "Salmon", 7504122
    AddColorItem "Salmon 1", 6917375
    AddColorItem "Salmon 2", 6456046
    AddColorItem "Salmon 3", 5533901
    AddColorItem "Salmon 4", 3755147
    AddColorItem "Sienna", 2970272
    AddColorItem "Sienna 1", 4686591
    AddColorItem "Sienna 2", 4356590
    AddColorItem "Sienna 3", 3762381
    AddColorItem "Sienna 4", 2508683
    AddColorItem "Deep Pink", 9639167
    AddColorItem "Deep Pink 2", 8983278
    AddColorItem "Deep Pink 3", 7737549
    AddColorItem "Deep Pink 4", 5245579
    AddColorItem "Hot Pink", 11823615
    AddColorItem "Hot Pink 1", 11824895
    AddColorItem "Hot Pink 2", 10971886
    AddColorItem "Hot Pink 3", 9461965
    AddColorItem "Hot Pink 4", 6437515
    AddColorItem "Indian Red", 6053069
    AddColorItem "Indian Red 1", 6974207
    AddColorItem "Indian Red 2", 6513646
    AddColorItem "Indian Red 3", 5592525
    AddColorItem "Indian Red 4", 3816075
    AddColorItem "Light Pink", 12695295
    AddColorItem "Light Pink 1", 12168959
    AddColorItem "Light Pink 2", 11379438
    AddColorItem "Light Pink 3", 9800909
    AddColorItem "Light Pink 4", 6643595
    AddColorItem "Medium Violet Red", 8721863
    AddColorItem "Misty Rose", 14804223
    AddColorItem "Misty Rose 2", 13817326
    AddColorItem "Misty Rose 3", 11909069
    AddColorItem "Misty Rose 4", 8093067
    AddColorItem "Orange Red", 17919
    AddColorItem "Orange Red 2", 16622
    AddColorItem "Orange Red 3", 14285
    AddColorItem "Orange Red 4", 9611
    AddColorItem "Pale Violet Red", 9662683
    AddColorItem "Pale Violet Red 1", 11240191
    AddColorItem "Pale Violet Red 2", 10451438
    AddColorItem "Pale Violet Red 3", 9005261
    AddColorItem "Pale Violet Red 4", 6113163
    AddColorItem "Violet Red", 9445584
    AddColorItem "Violet Red 1", 9846527
    AddColorItem "Violet Red 2", 9190126
    AddColorItem "Violet Red 3", 7877325
    AddColorItem "Violet Red 4", 5382795
    AddColorItem "Firebrick", 2237106
    AddColorItem "Firebrick 1", 3158271
    AddColorItem "Firebrick 2", 2895086
    AddColorItem "Firebrick 3", 2500301
    AddColorItem "Firebrick 4", 1710731
    AddColorItem "Pink", 13353215
    AddColorItem "Pink 1", 12957183
    AddColorItem "Pink 2", 12102126
    AddColorItem "Pink 3", 10392013
    AddColorItem "Pink 4", 7103371
    AddColorItem "Red", 255
    AddColorItem "Red 2", 238
    AddColorItem "Red 3", 205
    AddColorItem "Red 4", 139
    AddColorItem "Tomato", 4678655
    AddColorItem "Tomato 2", 4349166
    AddColorItem "Tomato 3", 3755981
    AddColorItem "Tomato 4", 2504331
    AddColorItem "Dark Orchid", 13382297
    AddColorItem "Dark Orchid 1", 16727743
    AddColorItem "Dark Orchid 2", 15612594
    AddColorItem "Dark Orchid 3", 13447834
    AddColorItem "Dark Orchid 4", 9118312
    AddColorItem "Dark Violet", 13828244
    AddColorItem "Lavender Blush", 16118015
    AddColorItem "Lavender Blush 2", 15065326
    AddColorItem "Lavender Blush 3", 12960205
    AddColorItem "Lavender Blush 4", 8815499
    AddColorItem "Medium Orchid", 13850042
    AddColorItem "Medium Orchid 1", 16738016
    AddColorItem "Medium Orchid 2", 15622097
    AddColorItem "Medium Orchid 3", 13456052
    AddColorItem "Medium Orchid4", 9123706
    AddColorItem "Medium Purple", 14381203
    AddColorItem "Medium Purple 1", 16745131
    AddColorItem "Medium Purple 2", 15628703
    AddColorItem "Medium Purple 3", 13461641
    AddColorItem "Medium Purple 4", 9127773
    AddColorItem "Lavender", 16443110
    AddColorItem "Magenta", 16711935
    AddColorItem "Magenta 2", 15597806
    AddColorItem "Magenta 3", 13435085
    AddColorItem "Magenta 4", 9109643
    AddColorItem "Maroon", 6303920
    AddColorItem "Maroon 1", 11744511
    AddColorItem "Maroon 2", 10957038
    AddColorItem "Maroon 3", 9447885
    AddColorItem "Maroon 4", 6429835
    AddColorItem "Orchid", 14053594
    AddColorItem "Orchid 1", 16417791
    AddColorItem "Orchid 2", 15301358
    AddColorItem "Orchid 3", 13199821
    AddColorItem "Orchid 4", 8996747
    AddColorItem "Plum", 14524637
    AddColorItem "Plum 1", 16759807
    AddColorItem "Plum 2", 15642350
    AddColorItem "Plum 3", 13473485
    AddColorItem "Plum 4", 9135755
    AddColorItem "Purple", 15736992
    AddColorItem "Purple 1", 16724123
    AddColorItem "Purple 2", 15608977
    AddColorItem "Purple 3", 13444733
    AddColorItem "Purple 4", 9116245
    AddColorItem "Thistle", 14204888
    AddColorItem "Thistle 1", 16769535
    AddColorItem "Thistle 2", 15651566
    AddColorItem "Thistle 3", 13481421
    AddColorItem "Thistle 4", 9141131
    AddColorItem "Violet", 15631086
    AddColorItem "Antique White", 14150650
    AddColorItem "Antique White 1", 14413823
    AddColorItem "Antique White 2", 13426670
    AddColorItem "Antique White 3", 11583693
    AddColorItem "Antique White 4", 7897995
    AddColorItem "Floral White", 15792895
    AddColorItem "Ghost White", 16775416
    AddColorItem "Navajo White", 11394815
    AddColorItem "Navajo White 2", 10604526
    AddColorItem "Navajo White 3", 9155533
    AddColorItem "Navajo White 4", 6191499
    AddColorItem "Old Lace", 15136253
    AddColorItem "Gainsboro", 14474460
    AddColorItem "Ivory", 15794175
    AddColorItem "Ivory 2", 14741230
    AddColorItem "Ivory 3", 12701133
    AddColorItem "Ivory 4", 8620939
    AddColorItem "Linen", 15134970
    AddColorItem "Seashell", 15660543
    AddColorItem "Seashell 2", 14607854
    AddColorItem "Seashell 3", 12568013
    AddColorItem "Seashell 4", 8554123
    AddColorItem "Snow", 16448255
    AddColorItem "Snow 2", 15329774
    AddColorItem "Snow 3", 13224397
    AddColorItem "Snow 4", 9013643
    AddColorItem "Wheat", 11788021
    AddColorItem "Wheat 1", 12249087
    AddColorItem "Wheat 2", 11458798
    AddColorItem "Wheat 3", 9878221
    AddColorItem "Wheat 4", 6717067
    AddColorItem "Blanched Almond", 13495295
    AddColorItem "Dark Goldenrod", 755384
    AddColorItem "Dark Goldenrod 1", 1030655
    AddColorItem "Dark Goldenrod 2", 962030
    AddColorItem "Dark Goldenrod 3", 824781
    AddColorItem "Dark Goldenrod 4", 550283
    AddColorItem "Lemon Chiffon", 13499135
    AddColorItem "Lemon Chiffon 2", 12577262
    AddColorItem "Lemon Chiffon 3", 10865101
    AddColorItem "Lemon Chiffon 4", 7375243
    AddColorItem "Light Goldenrod", 8576494
    AddColorItem "Light Goldenrod 1", 9170175
    AddColorItem "Light Goldenrod 2", 8576238
    AddColorItem "Light Goldenrod 3", 7388877
    AddColorItem "Light Goldenrod 4", 5013899
    AddColorItem "Light Goldenrod Yellow", 13826810
    AddColorItem "Light Yellow", 14745599
    AddColorItem "Light Yellow 2", 13758190
    AddColorItem "Light Yellow 3", 11849165
    AddColorItem "Light Yellow 4", 8031115
    AddColorItem "Pale Goldenrod", 11200750
    AddColorItem "Papaya Whip", 14020607
    AddColorItem "Cornsilk", 14481663
    AddColorItem "Cornsilk 2", 13494510
    AddColorItem "Cornsilk 3", 11651277
    AddColorItem "Cornsilk 4", 7899275
    AddColorItem "Gold", 55295
    AddColorItem "Gold 2", 51694
    AddColorItem "Gold 3", 44493
    AddColorItem "Gold 4", 30091
    AddColorItem "Goldenrod", 2139610
    AddColorItem "Goldenrod 1", 2474495
    AddColorItem "Goldenrod 2", 2274542
    AddColorItem "Goldenrod 3", 1940429
    AddColorItem "Goldenrod 4", 1337739
    AddColorItem "Moccasin", 11920639
    AddColorItem "Yellow", 65535
    AddColorItem "Yellow 2", 61166
    AddColorItem "Yellow 3", 52685
    AddColorItem "Yellow 4", 35723

End Sub
