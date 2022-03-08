VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sudany 1.5"
   ClientHeight    =   10140
   ClientLeft      =   2175
   ClientTop       =   735
   ClientWidth     =   10305
   ForeColor       =   &H80000005&
   Icon            =   "sud2 rnd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10140
   ScaleWidth      =   10305
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2325
      Index           =   2
      Left            =   435
      TabIndex        =   762
      Top             =   840
      Visible         =   0   'False
      Width           =   2115
   End
   Begin MSComDlg.CommonDialog CommonDlg 
      Left            =   660
      Top             =   525
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmr_rnd 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1680
      Top             =   2640
   End
   Begin VB.CommandButton cmd_rnd 
      Caption         =   "&Random !"
      Height          =   375
      Left            =   1515
      TabIndex        =   4
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmd_chk25 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1387
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk25 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1386
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk25 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1385
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk25 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1384
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk25 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1383
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk25 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1382
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk25 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1381
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk25 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1380
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk25 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1379
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk25 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1378
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk25 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1377
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk25 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1376
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk25 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1375
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk25 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1374
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk25 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1373
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk25 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1372
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk25 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1371
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk25 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1370
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk25 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1369
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk25 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1368
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk25 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1367
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk25 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1366
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk25 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1365
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk25 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1364
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1363
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1362
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1361
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1360
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1359
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1358
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1357
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1356
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1355
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1354
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1353
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1352
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1351
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1350
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1349
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1348
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1347
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1346
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1345
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1344
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1343
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1342
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1341
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1340
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1339
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1338
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1337
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1336
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1335
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1334
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1333
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1332
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1331
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1330
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1329
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1328
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1327
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1326
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1325
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1324
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1323
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1322
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1321
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1320
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1319
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1318
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1317
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1316
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1315
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1314
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1313
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1312
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1311
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1310
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1309
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1308
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1307
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1306
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1305
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1304
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1303
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1302
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1301
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1300
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1299
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1298
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1297
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1296
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1295
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1294
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1293
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1292
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1291
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1290
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1289
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1288
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1287
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1286
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1285
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1284
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1283
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1282
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1281
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1280
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1279
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1278
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1277
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1276
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1275
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1274
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1273
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1272
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1271
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1270
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1269
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1268
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1267
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1266
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1265
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1264
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1263
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1262
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1261
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1260
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1259
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1258
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1257
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1256
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1255
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1254
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1253
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1252
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1251
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1250
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1249
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1248
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1247
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1246
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1245
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1244
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1243
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1242
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1241
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1240
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1239
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1238
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1237
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1236
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1235
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1234
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1233
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1232
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1231
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1230
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1229
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1228
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1227
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1226
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1225
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1224
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1223
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1222
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1221
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1220
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1219
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1218
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1217
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1216
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1215
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1214
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1213
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1212
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1211
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1210
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1209
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1208
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1207
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1206
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1205
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1204
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1203
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1202
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1201
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1200
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1199
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1198
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1197
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1196
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1195
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1194
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1193
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1192
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1191
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1190
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1189
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1188
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1187
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1186
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1185
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1184
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1183
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1182
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1181
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1180
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1179
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1178
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1177
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1176
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1175
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1174
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1173
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1172
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1171
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1170
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1169
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1168
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1167
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1166
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1165
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1164
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1163
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1162
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1161
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1160
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1159
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1158
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1157
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1156
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1155
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1154
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1153
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1152
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1151
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1150
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1149
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1148
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1147
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1146
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1145
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1144
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1143
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1142
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1141
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1140
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1139
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1138
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1137
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1136
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1135
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1134
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1133
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1132
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1131
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1130
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1129
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1128
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1127
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1126
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1125
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1124
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1123
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1122
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1121
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1120
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1119
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1118
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1117
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1116
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1115
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1114
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1113
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1112
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1111
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1110
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1109
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1108
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1107
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1106
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1105
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1104
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1103
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1102
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1101
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1100
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1099
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1098
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1097
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1096
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1095
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1094
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1093
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1092
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1091
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1090
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1089
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1088
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1087
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1086
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1085
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1084
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1083
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1082
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1081
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1080
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1079
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1078
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1077
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1076
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1075
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1074
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1073
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1072
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1071
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1070
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1069
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1068
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1067
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1066
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1065
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1064
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1063
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1062
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1061
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1060
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1059
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1058
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1057
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1056
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1055
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1054
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1053
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1052
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1051
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1050
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1049
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1048
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1047
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1046
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1045
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1044
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1043
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1042
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1041
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1040
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1039
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1038
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1037
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1036
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1035
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1034
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1033
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1032
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1031
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1030
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1029
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1028
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1027
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1026
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1025
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1024
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1023
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1022
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1021
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1020
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1019
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1018
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1017
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1016
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1015
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1014
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1013
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1012
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1011
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1010
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1009
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1008
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1007
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1006
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1005
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1004
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1003
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1002
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1001
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   1000
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   999
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   998
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   997
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   996
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   995
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   994
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   993
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   992
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   991
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   990
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   989
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   988
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   987
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   986
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   985
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   984
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   983
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   982
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   981
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   980
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   979
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   978
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   977
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   976
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   975
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   974
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   973
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   972
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   971
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   970
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   969
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   968
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   967
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   966
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   965
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   964
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   963
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   962
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   961
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   960
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   959
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   958
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   957
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   956
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   955
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   954
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   953
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   952
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   951
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   950
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   949
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   948
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   947
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   946
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   945
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   944
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   943
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   942
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   941
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   940
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   939
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   938
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   937
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   936
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   935
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   934
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   933
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   932
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   931
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   930
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   929
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   928
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   927
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   926
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   925
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   924
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   923
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   922
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   921
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   920
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   919
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   918
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   917
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   916
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   915
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   914
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   913
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   912
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   911
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   910
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   909
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   908
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk21 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   8145
      Style           =   1  'Graphical
      TabIndex        =   907
      Top             =   3360
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk22 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   8385
      Style           =   1  'Graphical
      TabIndex        =   906
      Top             =   3360
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk23 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   8625
      Style           =   1  'Graphical
      TabIndex        =   905
      Top             =   3360
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk24 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   8865
      Style           =   1  'Graphical
      TabIndex        =   904
      Top             =   3360
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk25 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   9105
      Style           =   1  'Graphical
      TabIndex        =   903
      Top             =   3360
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk16 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   6945
      Style           =   1  'Graphical
      TabIndex        =   902
      Top             =   3360
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk17 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   7185
      Style           =   1  'Graphical
      TabIndex        =   901
      Top             =   3360
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk18 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   7425
      Style           =   1  'Graphical
      TabIndex        =   900
      Top             =   3360
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk19 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   7665
      Style           =   1  'Graphical
      TabIndex        =   899
      Top             =   3360
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk20 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   7905
      Style           =   1  'Graphical
      TabIndex        =   898
      Top             =   3360
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk11 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   5730
      Style           =   1  'Graphical
      TabIndex        =   897
      Top             =   3360
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk12 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   5970
      Style           =   1  'Graphical
      TabIndex        =   896
      Top             =   3360
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk13 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   6210
      Style           =   1  'Graphical
      TabIndex        =   895
      Top             =   3360
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk14 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   6450
      Style           =   1  'Graphical
      TabIndex        =   894
      Top             =   3360
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk15 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   6690
      Style           =   1  'Graphical
      TabIndex        =   893
      Top             =   3360
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk6 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   4515
      Style           =   1  'Graphical
      TabIndex        =   892
      Top             =   3360
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk7 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   4755
      Style           =   1  'Graphical
      TabIndex        =   891
      Top             =   3360
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk8 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   4995
      Style           =   1  'Graphical
      TabIndex        =   890
      Top             =   3360
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk9 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   5235
      Style           =   1  'Graphical
      TabIndex        =   889
      Top             =   3360
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk10 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   5475
      Style           =   1  'Graphical
      TabIndex        =   888
      Top             =   3360
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   887
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   886
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   885
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   884
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   883
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   882
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   881
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   880
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   879
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   878
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   877
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   876
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   875
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   874
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   873
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   872
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   871
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   870
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   869
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   868
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   867
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   866
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   865
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   864
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   863
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   862
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   861
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   860
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   859
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   858
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   857
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   856
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   855
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   854
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   853
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   852
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   851
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   850
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   849
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   848
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   847
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   846
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   845
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   844
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   843
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   842
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   841
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   840
      Top             =   5040
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   839
      Top             =   8160
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   838
      Top             =   8400
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   837
      Top             =   8640
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   836
      Top             =   8880
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   835
      Top             =   9120
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   834
      Top             =   6960
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   833
      Top             =   7200
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   832
      Top             =   7440
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   831
      Top             =   7680
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   830
      Top             =   7920
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   829
      Top             =   5760
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   828
      Top             =   6000
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   827
      Top             =   6240
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   826
      Top             =   6480
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   825
      Top             =   6720
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   824
      Top             =   5500
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   823
      Top             =   5260
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   822
      Top             =   5020
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   821
      Top             =   4800
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   820
      Top             =   4560
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   819
      Top             =   4320
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   818
      Top             =   4080
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   817
      Top             =   3840
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   816
      Top             =   3600
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   815
      Top             =   9120
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   814
      Top             =   8880
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   813
      Top             =   8640
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   812
      Top             =   8400
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   811
      Top             =   8160
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   810
      Top             =   7920
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   809
      Top             =   7680
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   808
      Top             =   7440
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   807
      Top             =   7200
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   806
      Top             =   6960
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   805
      Top             =   6720
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   804
      Top             =   6480
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   803
      Top             =   6240
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   802
      Top             =   6000
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   801
      Top             =   5760
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   800
      Top             =   5500
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   799
      Top             =   5260
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   798
      Top             =   5020
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   797
      Top             =   4800
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   796
      Top             =   4560
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   795
      Top             =   4320
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   794
      Top             =   4080
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   793
      Top             =   3840
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   792
      Top             =   3600
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk5 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   4275
      Style           =   1  'Graphical
      TabIndex        =   791
      Top             =   3360
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   4035
      Style           =   1  'Graphical
      TabIndex        =   790
      Top             =   3360
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk3 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   789
      Top             =   3360
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   788
      Top             =   3360
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   24
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   787
      Top             =   9120
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   786
      Top             =   8880
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   785
      Top             =   8640
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   784
      Top             =   8400
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   783
      Top             =   8160
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   782
      Top             =   7920
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   781
      Top             =   7680
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   780
      Top             =   7440
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   779
      Top             =   7200
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   778
      Top             =   6960
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   777
      Top             =   6720
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   776
      Top             =   6480
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   775
      Top             =   6240
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   774
      Top             =   6000
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   773
      Top             =   5760
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   772
      Top             =   5500
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   771
      Top             =   5260
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   770
      Top             =   5020
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   769
      Top             =   4800
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   768
      Top             =   4560
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   767
      Top             =   4320
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   766
      Top             =   4080
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   765
      Top             =   3840
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   764
      Top             =   3600
      Width           =   225
   End
   Begin VB.CommandButton cmd_chk1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   763
      Top             =   3360
      Width           =   225
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7680
      Index           =   1
      Left            =   9405
      TabIndex        =   761
      Top             =   2295
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   0
      Left            =   3075
      TabIndex        =   760
      Top             =   9450
      Visible         =   0   'False
      Width           =   6420
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Height          =   6135
      Left            =   3315
      TabIndex        =   758
      Top             =   3240
      Visible         =   0   'False
      Width           =   6020
      Begin VB.Line Line7 
         X1              =   0
         X2              =   6000
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         Index           =   25
         X1              =   -45
         X2              =   -45
         Y1              =   1455
         Y2              =   7470
      End
      Begin VB.Line Line30 
         BorderWidth     =   2
         X1              =   0
         X2              =   6500
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Line Line29 
         X1              =   0
         X2              =   6000
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Line Line28 
         X1              =   0
         X2              =   6000
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Line Line27 
         X1              =   0
         X2              =   6000
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Line Line26 
         X1              =   0
         X2              =   6000
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Line Line25 
         BorderWidth     =   2
         X1              =   0
         X2              =   6000
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Line Line24 
         X1              =   0
         X2              =   6000
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line23 
         X1              =   0
         X2              =   6000
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line22 
         X1              =   0
         X2              =   6000
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Line21 
         X1              =   0
         X2              =   6000
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line20 
         BorderWidth     =   2
         X1              =   0
         X2              =   6000
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line19 
         X1              =   0
         X2              =   6000
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line18 
         X1              =   0
         X2              =   6000
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line17 
         X1              =   0
         X2              =   6000
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line16 
         X1              =   0
         X2              =   6000
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line15 
         BorderWidth     =   2
         X1              =   0
         X2              =   6000
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line14 
         X1              =   0
         X2              =   6000
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line13 
         X1              =   0
         X2              =   6000
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line12 
         X1              =   0
         X2              =   6000
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line11 
         X1              =   0
         X2              =   6000
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line10 
         BorderWidth     =   2
         X1              =   0
         X2              =   6000
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line9 
         X1              =   0
         X2              =   6000
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line8 
         X1              =   0
         X2              =   6000
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   6000
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         Index           =   24
         X1              =   6000
         X2              =   6000
         Y1              =   120
         Y2              =   6120
      End
      Begin VB.Line Line5 
         Index           =   23
         X1              =   5760
         X2              =   5760
         Y1              =   120
         Y2              =   6120
      End
      Begin VB.Line Line5 
         Index           =   22
         X1              =   5520
         X2              =   5520
         Y1              =   120
         Y2              =   6120
      End
      Begin VB.Line Line5 
         Index           =   21
         X1              =   5280
         X2              =   5280
         Y1              =   120
         Y2              =   6120
      End
      Begin VB.Line Line5 
         Index           =   20
         X1              =   5040
         X2              =   5040
         Y1              =   120
         Y2              =   6120
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         Index           =   19
         X1              =   4800
         X2              =   4800
         Y1              =   0
         Y2              =   6120
      End
      Begin VB.Line Line5 
         Index           =   18
         X1              =   4560
         X2              =   4560
         Y1              =   120
         Y2              =   6120
      End
      Begin VB.Line Line5 
         Index           =   17
         X1              =   4320
         X2              =   4320
         Y1              =   120
         Y2              =   6120
      End
      Begin VB.Line Line5 
         Index           =   16
         X1              =   4080
         X2              =   4080
         Y1              =   120
         Y2              =   6120
      End
      Begin VB.Line Line5 
         Index           =   15
         X1              =   3840
         X2              =   3840
         Y1              =   120
         Y2              =   6120
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         Index           =   14
         X1              =   3600
         X2              =   3600
         Y1              =   0
         Y2              =   6120
      End
      Begin VB.Line Line5 
         Index           =   13
         X1              =   3360
         X2              =   3360
         Y1              =   120
         Y2              =   6120
      End
      Begin VB.Line Line5 
         Index           =   12
         X1              =   3120
         X2              =   3120
         Y1              =   120
         Y2              =   6120
      End
      Begin VB.Line Line5 
         Index           =   11
         X1              =   2880
         X2              =   2880
         Y1              =   120
         Y2              =   6120
      End
      Begin VB.Line Line5 
         Index           =   10
         X1              =   2640
         X2              =   2640
         Y1              =   120
         Y2              =   6120
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         Index           =   9
         X1              =   2400
         X2              =   2400
         Y1              =   0
         Y2              =   6120
      End
      Begin VB.Line Line5 
         Index           =   8
         X1              =   2160
         X2              =   2160
         Y1              =   120
         Y2              =   6120
      End
      Begin VB.Line Line5 
         Index           =   7
         X1              =   1920
         X2              =   1920
         Y1              =   120
         Y2              =   6120
      End
      Begin VB.Line Line5 
         Index           =   6
         X1              =   1680
         X2              =   1680
         Y1              =   120
         Y2              =   6120
      End
      Begin VB.Line Line5 
         Index           =   5
         X1              =   1440
         X2              =   1440
         Y1              =   120
         Y2              =   6120
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         Index           =   4
         X1              =   1200
         X2              =   1200
         Y1              =   0
         Y2              =   6120
      End
      Begin VB.Line Line5 
         Index           =   3
         X1              =   960
         X2              =   960
         Y1              =   120
         Y2              =   6120
      End
      Begin VB.Line Line5 
         Index           =   2
         X1              =   720
         X2              =   720
         Y1              =   120
         Y2              =   6120
      End
      Begin VB.Line Line5 
         Index           =   1
         X1              =   480
         X2              =   480
         Y1              =   120
         Y2              =   6120
      End
      Begin VB.Line Line5 
         Index           =   0
         X1              =   240
         X2              =   240
         Y1              =   120
         Y2              =   6120
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2625
      MultiLine       =   -1  'True
      TabIndex        =   759
      Top             =   2205
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.CommandButton cmd_N 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Negative"
      Height          =   375
      Left            =   555
      TabIndex        =   3
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   45
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   990
      TabIndex        =   6
      Top             =   45
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmd_print 
      Caption         =   "&PRINT"
      Height          =   375
      Left            =   1860
      TabIndex        =   7
      Top             =   45
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command50 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   3360
      MaskColor       =   &H0000FFFF&
      TabIndex        =   757
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command49 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   3600
      MaskColor       =   &H0000FFFF&
      TabIndex        =   756
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command48 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   3840
      MaskColor       =   &H0000FFFF&
      TabIndex        =   755
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command47 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   4080
      MaskColor       =   &H0000FFFF&
      TabIndex        =   754
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command46 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   4320
      MaskColor       =   &H0000FFFF&
      TabIndex        =   753
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command45 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   4545
      MaskColor       =   &H0000FFFF&
      TabIndex        =   752
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command44 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   4785
      MaskColor       =   &H0000FFFF&
      TabIndex        =   751
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command43 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   5025
      MaskColor       =   &H0000FFFF&
      TabIndex        =   750
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command42 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   5265
      MaskColor       =   &H0000FFFF&
      TabIndex        =   749
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command41 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   5505
      MaskColor       =   &H0000FFFF&
      TabIndex        =   748
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command40 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   5745
      MaskColor       =   &H0000FFFF&
      TabIndex        =   747
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command39 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   5985
      MaskColor       =   &H0000FFFF&
      TabIndex        =   746
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command38 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   6225
      MaskColor       =   &H0000FFFF&
      TabIndex        =   745
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command37 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   6465
      MaskColor       =   &H0000FFFF&
      TabIndex        =   744
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command36 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   6705
      MaskColor       =   &H0000FFFF&
      TabIndex        =   743
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command35 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   6945
      MaskColor       =   &H0000FFFF&
      TabIndex        =   742
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command34 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   7185
      MaskColor       =   &H0000FFFF&
      TabIndex        =   741
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command33 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   7425
      MaskColor       =   &H0000FFFF&
      TabIndex        =   740
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command32 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   7665
      MaskColor       =   &H0000FFFF&
      TabIndex        =   739
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command31 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   7905
      MaskColor       =   &H0000FFFF&
      TabIndex        =   738
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command30 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   8160
      MaskColor       =   &H0000FFFF&
      TabIndex        =   737
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command29 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   8400
      MaskColor       =   &H0000FFFF&
      TabIndex        =   736
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command28 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   8640
      MaskColor       =   &H0000FFFF&
      TabIndex        =   735
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command27 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   8880
      MaskColor       =   &H0000FFFF&
      TabIndex        =   734
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command26 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Left            =   9120
      MaskColor       =   &H0000FFFF&
      TabIndex        =   733
      ToolTipText     =   "clears all tur!"
      Top             =   9840
      Width           =   160
   End
   Begin VB.CommandButton Command25 
      Caption         =   "V"
      Height          =   255
      Left            =   9120
      MaskColor       =   &H0000FFFF&
      TabIndex        =   732
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton Command24 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Left            =   8880
      MaskColor       =   &H0000FFFF&
      TabIndex        =   731
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Left            =   8640
      MaskColor       =   &H0000FFFF&
      TabIndex        =   730
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Left            =   8400
      MaskColor       =   &H0000FFFF&
      TabIndex        =   729
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Left            =   8160
      MaskColor       =   &H0000FFFF&
      TabIndex        =   728
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Left            =   7905
      MaskColor       =   &H0000FFFF&
      TabIndex        =   727
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Left            =   7665
      MaskColor       =   &H0000FFFF&
      TabIndex        =   726
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Left            =   7425
      MaskColor       =   &H0000FFFF&
      TabIndex        =   725
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Left            =   7185
      MaskColor       =   &H0000FFFF&
      TabIndex        =   724
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Left            =   6945
      MaskColor       =   &H0000FFFF&
      TabIndex        =   723
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Left            =   6705
      MaskColor       =   &H0000FFFF&
      TabIndex        =   722
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Left            =   6465
      MaskColor       =   &H0000FFFF&
      TabIndex        =   721
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Left            =   6225
      MaskColor       =   &H0000FFFF&
      TabIndex        =   720
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Left            =   5985
      MaskColor       =   &H0000FFFF&
      TabIndex        =   719
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Left            =   5745
      MaskColor       =   &H0000FFFF&
      TabIndex        =   718
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Left            =   5505
      MaskColor       =   &H0000FFFF&
      TabIndex        =   717
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Left            =   5265
      MaskColor       =   &H0000FFFF&
      TabIndex        =   716
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Left            =   5025
      MaskColor       =   &H0000FFFF&
      TabIndex        =   715
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Left            =   4785
      MaskColor       =   &H0000FFFF&
      TabIndex        =   714
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Left            =   4545
      MaskColor       =   &H0000FFFF&
      TabIndex        =   713
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Left            =   4320
      MaskColor       =   &H0000FFFF&
      TabIndex        =   712
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Left            =   4080
      MaskColor       =   &H0000FFFF&
      TabIndex        =   711
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Left            =   3840
      MaskColor       =   &H0000FFFF&
      TabIndex        =   710
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Left            =   3600
      MaskColor       =   &H0000FFFF&
      TabIndex        =   709
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Left            =   3360
      MaskColor       =   &H0000FFFF&
      TabIndex        =   708
      ToolTipText     =   "mark all tur!"
      Top             =   9480
      Width           =   160
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   24
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   707
      ToolTipText     =   "clears all line!"
      Top             =   9120
      Width           =   220
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   23
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   706
      ToolTipText     =   "clears all line!"
      Top             =   8880
      Width           =   220
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   22
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   705
      ToolTipText     =   "clears all line!"
      Top             =   8640
      Width           =   220
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   21
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   704
      ToolTipText     =   "clears all line!"
      Top             =   8400
      Width           =   220
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   20
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   703
      ToolTipText     =   "clears all line!"
      Top             =   8160
      Width           =   220
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   19
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   702
      ToolTipText     =   "clears all line!"
      Top             =   7920
      Width           =   220
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   18
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   701
      ToolTipText     =   "clears all line!"
      Top             =   7680
      Width           =   220
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   17
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   700
      ToolTipText     =   "clears all line!"
      Top             =   7440
      Width           =   220
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   16
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   699
      ToolTipText     =   "clears all line!"
      Top             =   7200
      Width           =   220
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   15
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   698
      ToolTipText     =   "clears all line!"
      Top             =   6960
      Width           =   220
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   14
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   697
      ToolTipText     =   "clears all line!"
      Top             =   6720
      Width           =   220
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   13
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   696
      ToolTipText     =   "clears all line!"
      Top             =   6480
      Width           =   220
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   12
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   695
      ToolTipText     =   "clears all line!"
      Top             =   6240
      Width           =   220
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   11
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   694
      ToolTipText     =   "clears all line!"
      Top             =   6000
      Width           =   220
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   10
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   693
      ToolTipText     =   "clears all line!"
      Top             =   5760
      Width           =   220
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   9
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   692
      ToolTipText     =   "clears all line!"
      Top             =   5520
      Width           =   220
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   8
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   691
      ToolTipText     =   "clears all line!"
      Top             =   5280
      Width           =   220
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   7
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   690
      ToolTipText     =   "clears all line!"
      Top             =   5040
      Width           =   220
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   6
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   689
      ToolTipText     =   "clears all line!"
      Top             =   4800
      Width           =   220
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   5
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   688
      ToolTipText     =   "clears all line!"
      Top             =   4560
      Width           =   220
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   4
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   687
      ToolTipText     =   "clears all line!"
      Top             =   4320
      Width           =   220
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   3
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   686
      ToolTipText     =   "clears all line!"
      Top             =   4080
      Width           =   220
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   2
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   685
      ToolTipText     =   "clears all line!"
      Top             =   3840
      Width           =   220
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   1
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   684
      ToolTipText     =   "clears all line!"
      Top             =   3600
      Width           =   220
   End
   Begin VB.CommandButton cmd_clr_line 
      BackColor       =   &H80000013&
      Caption         =   "X"
      Height          =   255
      Index           =   0
      Left            =   9915
      MaskColor       =   &H80000013&
      TabIndex        =   683
      ToolTipText     =   "clears all line!"
      Top             =   3360
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Index           =   24
      Left            =   9555
      MaskColor       =   &H80000013&
      TabIndex        =   682
      ToolTipText     =   "marks all line!"
      Top             =   9120
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Index           =   23
      Left            =   9555
      MaskColor       =   &H80000013&
      TabIndex        =   681
      ToolTipText     =   "marks all line!"
      Top             =   8880
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Index           =   22
      Left            =   9555
      MaskColor       =   &H80000013&
      TabIndex        =   680
      ToolTipText     =   "marks all line!"
      Top             =   8640
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Index           =   21
      Left            =   9555
      MaskColor       =   &H80000013&
      TabIndex        =   679
      ToolTipText     =   "marks all line!"
      Top             =   8400
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   250
      Index           =   20
      Left            =   9555
      MaskColor       =   &H80000013&
      TabIndex        =   678
      ToolTipText     =   "marks all line!"
      Top             =   8160
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Index           =   19
      Left            =   9555
      MaskColor       =   &H80000013&
      TabIndex        =   677
      ToolTipText     =   "marks all line!"
      Top             =   7920
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Index           =   18
      Left            =   9555
      MaskColor       =   &H80000013&
      TabIndex        =   676
      ToolTipText     =   "marks all line!"
      Top             =   7680
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Index           =   17
      Left            =   9555
      MaskColor       =   &H80000013&
      TabIndex        =   675
      ToolTipText     =   "marks all line!"
      Top             =   7440
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Index           =   16
      Left            =   9555
      MaskColor       =   &H80000013&
      TabIndex        =   674
      ToolTipText     =   "marks all line!"
      Top             =   7200
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Index           =   15
      Left            =   9555
      MaskColor       =   &H80000013&
      TabIndex        =   673
      ToolTipText     =   "marks all line!"
      Top             =   6960
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Index           =   14
      Left            =   9555
      MaskColor       =   &H80000013&
      TabIndex        =   672
      ToolTipText     =   "marks all line!"
      Top             =   6720
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Index           =   13
      Left            =   9555
      MaskColor       =   &H80000013&
      TabIndex        =   671
      ToolTipText     =   "marks all line!"
      Top             =   6480
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Index           =   12
      Left            =   9555
      MaskColor       =   &H80000013&
      TabIndex        =   670
      ToolTipText     =   "marks all line!"
      Top             =   6240
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Index           =   11
      Left            =   9555
      MaskColor       =   &H80000013&
      TabIndex        =   669
      ToolTipText     =   "marks all line!"
      Top             =   6000
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Index           =   10
      Left            =   9555
      MaskColor       =   &H80000013&
      TabIndex        =   668
      ToolTipText     =   "marks all line!"
      Top             =   5760
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Index           =   9
      Left            =   9555
      MaskColor       =   &H80000013&
      TabIndex        =   667
      ToolTipText     =   "marks all line!"
      Top             =   5520
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Index           =   8
      Left            =   9555
      MaskColor       =   &H80000013&
      TabIndex        =   666
      ToolTipText     =   "marks all line!"
      Top             =   5280
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Index           =   7
      Left            =   9555
      MaskColor       =   &H80000013&
      TabIndex        =   665
      ToolTipText     =   "marks all line!"
      Top             =   5040
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Index           =   6
      Left            =   9555
      MaskColor       =   &H80000013&
      TabIndex        =   664
      ToolTipText     =   "marks all line!"
      Top             =   4800
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Index           =   5
      Left            =   9555
      MaskColor       =   &H80000013&
      TabIndex        =   663
      ToolTipText     =   "marks all line!"
      Top             =   4560
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Index           =   4
      Left            =   9555
      MaskColor       =   &H80000013&
      TabIndex        =   662
      ToolTipText     =   "marks all line!"
      Top             =   4320
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Index           =   3
      Left            =   9555
      MaskColor       =   &H80000013&
      TabIndex        =   661
      ToolTipText     =   "marks all line!"
      Top             =   4080
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Index           =   2
      Left            =   9555
      MaskColor       =   &H80000013&
      TabIndex        =   660
      ToolTipText     =   "marks all line!"
      Top             =   3840
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H80000013&
      Caption         =   "V"
      Height          =   255
      Index           =   1
      Left            =   9555
      MaskColor       =   &H80000013&
      TabIndex        =   659
      ToolTipText     =   "marks all line!"
      Top             =   3600
      Width           =   220
   End
   Begin VB.CommandButton cmd_mrk_line 
      BackColor       =   &H8000000D&
      Caption         =   "V"
      Height          =   255
      Index           =   0
      Left            =   9555
      MaskColor       =   &H0000FFFF&
      TabIndex        =   658
      ToolTipText     =   "marks all line!"
      Top             =   3360
      Width           =   220
   End
   Begin VB.CommandButton cmd_calc 
      Caption         =   "C&ALC !"
      Height          =   375
      Left            =   1050
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmd_cln 
      Caption         =   "&Clean All"
      Height          =   375
      Left            =   1515
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "delete all of the buttons!"
      Top             =   1650
      Width           =   855
   End
   Begin VB.CommandButton cmd_mark 
      Caption         =   "&Mark All"
      Height          =   375
      Left            =   555
      TabIndex        =   1
      ToolTipText     =   "Marks all of the buttons!"
      Top             =   1650
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   2
      X1              =   75
      X2              =   10370
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   3
      X1              =   75
      X2              =   10370
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   75
      X2              =   10370
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   75
      X2              =   10370
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line31 
      X1              =   1965
      X2              =   1965
      Y1              =   3360
      Y2              =   9370
   End
   Begin VB.Line Line32 
      X1              =   765
      X2              =   765
      Y1              =   3360
      Y2              =   9370
   End
   Begin VB.Line Line34 
      X1              =   3315
      X2              =   9335
      Y1              =   825
      Y2              =   825
   End
   Begin VB.Line Line33 
      X1              =   3315
      X2              =   9335
      Y1              =   2025
      Y2              =   2025
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      DrawMode        =   1  'Blackness
      Index           =   4
      X1              =   8115
      X2              =   8115
      Y1              =   3360
      Y2              =   10440
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      DrawMode        =   1  'Blackness
      Index           =   3
      X1              =   6915
      X2              =   6915
      Y1              =   3360
      Y2              =   10440
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      DrawMode        =   1  'Blackness
      Index           =   2
      X1              =   5715
      X2              =   5715
      Y1              =   3360
      Y2              =   10440
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      DrawMode        =   1  'Blackness
      Index           =   1
      X1              =   4515
      X2              =   4515
      Y1              =   3375
      Y2              =   10455
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   8115
      X2              =   8115
      Y1              =   120
      Y2              =   10200
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      Index           =   1
      X1              =   6915
      X2              =   6915
      Y1              =   120
      Y2              =   10200
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      Index           =   0
      X1              =   5715
      X2              =   5715
      Y1              =   120
      Y2              =   10200
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      DrawMode        =   1  'Blackness
      Index           =   0
      X1              =   4515
      X2              =   4515
      Y1              =   120
      Y2              =   10200
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   3315
      TabIndex        =   632
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   3555
      TabIndex        =   631
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   3795
      TabIndex        =   630
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   4035
      TabIndex        =   629
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   4275
      TabIndex        =   628
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   4515
      TabIndex        =   627
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   4755
      TabIndex        =   626
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   4995
      TabIndex        =   625
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   5235
      TabIndex        =   624
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   5475
      TabIndex        =   623
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   5715
      TabIndex        =   622
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   5955
      TabIndex        =   621
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   6195
      TabIndex        =   620
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   6435
      TabIndex        =   619
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   6675
      TabIndex        =   618
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   6915
      TabIndex        =   617
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   7155
      TabIndex        =   616
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   7395
      TabIndex        =   615
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   7635
      TabIndex        =   614
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   7875
      TabIndex        =   613
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   8115
      TabIndex        =   612
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   8355
      TabIndex        =   611
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   8595
      TabIndex        =   610
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   8835
      TabIndex        =   609
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   9075
      TabIndex        =   608
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   9075
      TabIndex        =   391
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   8835
      TabIndex        =   390
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   8595
      TabIndex        =   389
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   8355
      TabIndex        =   388
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   8115
      TabIndex        =   387
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   9075
      TabIndex        =   386
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   8835
      TabIndex        =   385
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   8595
      TabIndex        =   384
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   8355
      TabIndex        =   383
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   8115
      TabIndex        =   382
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   9075
      TabIndex        =   381
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   8835
      TabIndex        =   380
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   8595
      TabIndex        =   379
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   8355
      TabIndex        =   378
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   8115
      TabIndex        =   377
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   9075
      TabIndex        =   376
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   8835
      TabIndex        =   375
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   8595
      TabIndex        =   374
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   8355
      TabIndex        =   373
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   8115
      TabIndex        =   372
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   9075
      TabIndex        =   371
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   8835
      TabIndex        =   370
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   8595
      TabIndex        =   369
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   8355
      TabIndex        =   368
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   8115
      TabIndex        =   367
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   9075
      TabIndex        =   366
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   8835
      TabIndex        =   365
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   8595
      TabIndex        =   364
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   8355
      TabIndex        =   363
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   8115
      TabIndex        =   362
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   9075
      TabIndex        =   361
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   8835
      TabIndex        =   360
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   8595
      TabIndex        =   359
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   8355
      TabIndex        =   358
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   8115
      TabIndex        =   357
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   9075
      TabIndex        =   356
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   8835
      TabIndex        =   355
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   8595
      TabIndex        =   354
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   8355
      TabIndex        =   353
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   8115
      TabIndex        =   352
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   9075
      TabIndex        =   351
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   8835
      TabIndex        =   350
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   8595
      TabIndex        =   349
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   8355
      TabIndex        =   348
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   8115
      TabIndex        =   347
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   9075
      TabIndex        =   346
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   8835
      TabIndex        =   345
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   8595
      TabIndex        =   344
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   8355
      TabIndex        =   343
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   8115
      TabIndex        =   342
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   9075
      TabIndex        =   341
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   8835
      TabIndex        =   340
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   8595
      TabIndex        =   339
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   8355
      TabIndex        =   338
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   8115
      TabIndex        =   337
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   9075
      TabIndex        =   336
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   8835
      TabIndex        =   335
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   8595
      TabIndex        =   334
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   8355
      TabIndex        =   333
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   8115
      TabIndex        =   332
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   7875
      TabIndex        =   331
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   7635
      TabIndex        =   330
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   7875
      TabIndex        =   329
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   7635
      TabIndex        =   328
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   7875
      TabIndex        =   327
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   7635
      TabIndex        =   326
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   7875
      TabIndex        =   325
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   7635
      TabIndex        =   324
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   7875
      TabIndex        =   323
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   7635
      TabIndex        =   322
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   7875
      TabIndex        =   321
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   7635
      TabIndex        =   320
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   7875
      TabIndex        =   319
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   7635
      TabIndex        =   318
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   7875
      TabIndex        =   317
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   7635
      TabIndex        =   316
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   7875
      TabIndex        =   315
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   7635
      TabIndex        =   314
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   7875
      TabIndex        =   313
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   7635
      TabIndex        =   312
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   7875
      TabIndex        =   311
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   7635
      TabIndex        =   310
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   7875
      TabIndex        =   309
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   7635
      TabIndex        =   308
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   7395
      TabIndex        =   307
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   7395
      TabIndex        =   306
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   7395
      TabIndex        =   305
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   7395
      TabIndex        =   304
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   7395
      TabIndex        =   303
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   7395
      TabIndex        =   302
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   7395
      TabIndex        =   301
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   7395
      TabIndex        =   300
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   7395
      TabIndex        =   299
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   7395
      TabIndex        =   298
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   7395
      TabIndex        =   297
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   7395
      TabIndex        =   296
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   7155
      TabIndex        =   295
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   7155
      TabIndex        =   294
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   7155
      TabIndex        =   293
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   7155
      TabIndex        =   292
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   7155
      TabIndex        =   291
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   7155
      TabIndex        =   290
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   7155
      TabIndex        =   289
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   7155
      TabIndex        =   288
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   7155
      TabIndex        =   287
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   7155
      TabIndex        =   286
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   7155
      TabIndex        =   285
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   7155
      TabIndex        =   284
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   6915
      TabIndex        =   283
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   6915
      TabIndex        =   282
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   6915
      TabIndex        =   281
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   6915
      TabIndex        =   280
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   6915
      TabIndex        =   279
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   6915
      TabIndex        =   278
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   6915
      TabIndex        =   277
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   6915
      TabIndex        =   276
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   6915
      TabIndex        =   275
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   6915
      TabIndex        =   274
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   6915
      TabIndex        =   273
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   6915
      TabIndex        =   272
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   6675
      TabIndex        =   271
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   6675
      TabIndex        =   270
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   6675
      TabIndex        =   269
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   6675
      TabIndex        =   268
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   6675
      TabIndex        =   267
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   6675
      TabIndex        =   266
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   6675
      TabIndex        =   265
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   6675
      TabIndex        =   264
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   6675
      TabIndex        =   263
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   6675
      TabIndex        =   262
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   6675
      TabIndex        =   261
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   6675
      TabIndex        =   260
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   6435
      TabIndex        =   259
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   6435
      TabIndex        =   258
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   6435
      TabIndex        =   257
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   6435
      TabIndex        =   256
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   6435
      TabIndex        =   255
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   6435
      TabIndex        =   254
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   6435
      TabIndex        =   253
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   6435
      TabIndex        =   252
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   6435
      TabIndex        =   251
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   6435
      TabIndex        =   250
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   6435
      TabIndex        =   249
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   6435
      TabIndex        =   248
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   6195
      TabIndex        =   247
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   6195
      TabIndex        =   246
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   6195
      TabIndex        =   245
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   6195
      TabIndex        =   244
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   6195
      TabIndex        =   243
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   6195
      TabIndex        =   242
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   6195
      TabIndex        =   241
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   6195
      TabIndex        =   240
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   6195
      TabIndex        =   239
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   6195
      TabIndex        =   238
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   6195
      TabIndex        =   237
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   6195
      TabIndex        =   236
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   5955
      TabIndex        =   235
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   5955
      TabIndex        =   234
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   5955
      TabIndex        =   233
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   5955
      TabIndex        =   232
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   5955
      TabIndex        =   231
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   5955
      TabIndex        =   230
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   5955
      TabIndex        =   229
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   5955
      TabIndex        =   228
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   5955
      TabIndex        =   227
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   5955
      TabIndex        =   226
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   5955
      TabIndex        =   225
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   5955
      TabIndex        =   224
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   5715
      TabIndex        =   223
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   5715
      TabIndex        =   222
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   5715
      TabIndex        =   221
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   5715
      TabIndex        =   220
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   5715
      TabIndex        =   219
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   5715
      TabIndex        =   218
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   5715
      TabIndex        =   217
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   5715
      TabIndex        =   216
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   5715
      TabIndex        =   215
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   5715
      TabIndex        =   214
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   5715
      TabIndex        =   213
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   5715
      TabIndex        =   212
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   5475
      TabIndex        =   187
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   5475
      TabIndex        =   186
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   5475
      TabIndex        =   185
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   5475
      TabIndex        =   184
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   5475
      TabIndex        =   183
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   5475
      TabIndex        =   182
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   5475
      TabIndex        =   181
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   5475
      TabIndex        =   180
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   5475
      TabIndex        =   179
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   5475
      TabIndex        =   178
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   5475
      TabIndex        =   177
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   5475
      TabIndex        =   176
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   5235
      TabIndex        =   175
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   5235
      TabIndex        =   174
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   5235
      TabIndex        =   173
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   5235
      TabIndex        =   172
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   5235
      TabIndex        =   171
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   5235
      TabIndex        =   170
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   5235
      TabIndex        =   169
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   5235
      TabIndex        =   168
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   5235
      TabIndex        =   167
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   5235
      TabIndex        =   166
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   5235
      TabIndex        =   165
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   5235
      TabIndex        =   164
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   4995
      TabIndex        =   163
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   4995
      TabIndex        =   162
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   4995
      TabIndex        =   161
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   4995
      TabIndex        =   160
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   4995
      TabIndex        =   159
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   4995
      TabIndex        =   158
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   4995
      TabIndex        =   157
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   4995
      TabIndex        =   156
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   4995
      TabIndex        =   155
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   4995
      TabIndex        =   154
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   4995
      TabIndex        =   153
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   4995
      TabIndex        =   152
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   4755
      TabIndex        =   151
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   4755
      TabIndex        =   150
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   4755
      TabIndex        =   149
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   4755
      TabIndex        =   148
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   4755
      TabIndex        =   147
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   4755
      TabIndex        =   146
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   4755
      TabIndex        =   145
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   4755
      TabIndex        =   144
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   4755
      TabIndex        =   143
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   4755
      TabIndex        =   142
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   4755
      TabIndex        =   141
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   4755
      TabIndex        =   140
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   4515
      TabIndex        =   139
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   4515
      TabIndex        =   138
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   4515
      TabIndex        =   137
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   4515
      TabIndex        =   136
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   4515
      TabIndex        =   135
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   4515
      TabIndex        =   134
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   4515
      TabIndex        =   133
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   4515
      TabIndex        =   132
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   4515
      TabIndex        =   131
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   4515
      TabIndex        =   130
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   4515
      TabIndex        =   129
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   4515
      TabIndex        =   128
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   4275
      TabIndex        =   127
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   4275
      TabIndex        =   126
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   4275
      TabIndex        =   125
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   4275
      TabIndex        =   124
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   4275
      TabIndex        =   123
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   4275
      TabIndex        =   122
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   4275
      TabIndex        =   121
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   4275
      TabIndex        =   120
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   4275
      TabIndex        =   119
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   4275
      TabIndex        =   118
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   4275
      TabIndex        =   117
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   4275
      TabIndex        =   116
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   4035
      TabIndex        =   115
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   4035
      TabIndex        =   114
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   4035
      TabIndex        =   113
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   4035
      TabIndex        =   112
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   4035
      TabIndex        =   111
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   4035
      TabIndex        =   110
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   4035
      TabIndex        =   109
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   4035
      TabIndex        =   108
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   4035
      TabIndex        =   107
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   4035
      TabIndex        =   106
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   4035
      TabIndex        =   105
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   4035
      TabIndex        =   104
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   3795
      TabIndex        =   103
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   3795
      TabIndex        =   102
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   3795
      TabIndex        =   101
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   3795
      TabIndex        =   100
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   3795
      TabIndex        =   99
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   3795
      TabIndex        =   98
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   3795
      TabIndex        =   97
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   3795
      TabIndex        =   96
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   3795
      TabIndex        =   95
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   3795
      TabIndex        =   94
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   3795
      TabIndex        =   93
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   3795
      TabIndex        =   92
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   3555
      TabIndex        =   31
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   3555
      TabIndex        =   30
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   3555
      TabIndex        =   29
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   3555
      TabIndex        =   28
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   3555
      TabIndex        =   27
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   3555
      TabIndex        =   26
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   3555
      TabIndex        =   25
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   3555
      TabIndex        =   24
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   3555
      TabIndex        =   23
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   3555
      TabIndex        =   22
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   3555
      TabIndex        =   21
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   3555
      TabIndex        =   20
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   3315
      TabIndex        =   19
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   3315
      TabIndex        =   18
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   3315
      TabIndex        =   17
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   3315
      TabIndex        =   16
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   3315
      TabIndex        =   15
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   3315
      TabIndex        =   14
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   3315
      TabIndex        =   13
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   3315
      TabIndex        =   12
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   3315
      TabIndex        =   11
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   3315
      TabIndex        =   10
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   3315
      TabIndex        =   9
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   3315
      TabIndex        =   8
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   324
      Left            =   2970
      TabIndex        =   657
      Top             =   9120
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   323
      Left            =   2730
      TabIndex        =   656
      Top             =   9120
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   322
      Left            =   2490
      TabIndex        =   655
      Top             =   9120
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   321
      Left            =   2250
      TabIndex        =   654
      Top             =   9120
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   320
      Left            =   2010
      TabIndex        =   653
      Top             =   9120
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   319
      Left            =   1770
      TabIndex        =   652
      Top             =   9120
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   318
      Left            =   1530
      TabIndex        =   651
      Top             =   9120
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   317
      Left            =   1290
      TabIndex        =   650
      Top             =   9120
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   316
      Left            =   1050
      TabIndex        =   649
      Top             =   9120
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   315
      Left            =   810
      TabIndex        =   648
      Top             =   9120
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   314
      Left            =   570
      TabIndex        =   647
      Top             =   9120
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   313
      Left            =   330
      TabIndex        =   646
      Top             =   9120
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   312
      Left            =   90
      TabIndex        =   645
      Top             =   9120
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   311
      Left            =   2970
      TabIndex        =   644
      Top             =   8880
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   310
      Left            =   2730
      TabIndex        =   643
      Top             =   8880
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   309
      Left            =   2490
      TabIndex        =   642
      Top             =   8880
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   308
      Left            =   2250
      TabIndex        =   641
      Top             =   8880
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   307
      Left            =   2010
      TabIndex        =   640
      Top             =   8880
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   306
      Left            =   1770
      TabIndex        =   639
      Top             =   8880
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   305
      Left            =   1530
      TabIndex        =   638
      Top             =   8880
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   304
      Left            =   1290
      TabIndex        =   637
      Top             =   8880
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   303
      Left            =   1050
      TabIndex        =   636
      Top             =   8880
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   302
      Left            =   810
      TabIndex        =   635
      Top             =   8880
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   301
      Left            =   570
      TabIndex        =   634
      Top             =   8880
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   300
      Left            =   330
      TabIndex        =   633
      Top             =   8880
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   299
      Left            =   90
      TabIndex        =   607
      Top             =   8880
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   298
      Left            =   2970
      TabIndex        =   606
      Top             =   8640
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   297
      Left            =   2730
      TabIndex        =   605
      Top             =   8640
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   296
      Left            =   2490
      TabIndex        =   604
      Top             =   8640
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   295
      Left            =   2250
      TabIndex        =   603
      Top             =   8640
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   294
      Left            =   2010
      TabIndex        =   602
      Top             =   8640
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   293
      Left            =   1770
      TabIndex        =   601
      Top             =   8640
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   292
      Left            =   1530
      TabIndex        =   600
      Top             =   8640
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   291
      Left            =   1290
      TabIndex        =   599
      Top             =   8640
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   290
      Left            =   1050
      TabIndex        =   598
      Top             =   8640
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   289
      Left            =   810
      TabIndex        =   597
      Top             =   8640
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   288
      Left            =   570
      TabIndex        =   596
      Top             =   8640
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   287
      Left            =   330
      TabIndex        =   595
      Top             =   8640
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   286
      Left            =   90
      TabIndex        =   594
      Top             =   8640
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   285
      Left            =   2970
      TabIndex        =   593
      Top             =   8400
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   284
      Left            =   2730
      TabIndex        =   592
      Top             =   8400
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   283
      Left            =   2490
      TabIndex        =   591
      Top             =   8400
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   282
      Left            =   2250
      TabIndex        =   590
      Top             =   8400
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   281
      Left            =   2010
      TabIndex        =   589
      Top             =   8400
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   280
      Left            =   1770
      TabIndex        =   588
      Top             =   8400
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   279
      Left            =   1530
      TabIndex        =   587
      Top             =   8400
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   278
      Left            =   1290
      TabIndex        =   586
      Top             =   8400
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   277
      Left            =   1050
      TabIndex        =   585
      Top             =   8400
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   276
      Left            =   810
      TabIndex        =   584
      Top             =   8400
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   275
      Left            =   570
      TabIndex        =   583
      Top             =   8400
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   274
      Left            =   330
      TabIndex        =   582
      Top             =   8400
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   273
      Left            =   90
      TabIndex        =   581
      Top             =   8400
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   272
      Left            =   2970
      TabIndex        =   580
      Top             =   8160
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   271
      Left            =   2730
      TabIndex        =   579
      Top             =   8160
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   270
      Left            =   2490
      TabIndex        =   578
      Top             =   8160
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   269
      Left            =   2250
      TabIndex        =   577
      Top             =   8160
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   268
      Left            =   2010
      TabIndex        =   576
      Top             =   8160
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   267
      Left            =   1770
      TabIndex        =   575
      Top             =   8160
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   266
      Left            =   1530
      TabIndex        =   574
      Top             =   8160
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   265
      Left            =   1290
      TabIndex        =   573
      Top             =   8160
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   264
      Left            =   1050
      TabIndex        =   572
      Top             =   8160
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   263
      Left            =   810
      TabIndex        =   571
      Top             =   8160
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   262
      Left            =   570
      TabIndex        =   570
      Top             =   8160
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   261
      Left            =   330
      TabIndex        =   569
      Top             =   8160
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   260
      Left            =   90
      TabIndex        =   568
      Top             =   8160
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   259
      Left            =   2970
      TabIndex        =   567
      Top             =   7920
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   258
      Left            =   2730
      TabIndex        =   566
      Top             =   7920
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   257
      Left            =   2490
      TabIndex        =   565
      Top             =   7920
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   256
      Left            =   2250
      TabIndex        =   564
      Top             =   7920
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   255
      Left            =   2010
      TabIndex        =   563
      Top             =   7920
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   254
      Left            =   1770
      TabIndex        =   562
      Top             =   7920
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   253
      Left            =   1530
      TabIndex        =   561
      Top             =   7920
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   252
      Left            =   1290
      TabIndex        =   560
      Top             =   7920
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   251
      Left            =   1050
      TabIndex        =   559
      Top             =   7920
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   250
      Left            =   810
      TabIndex        =   558
      Top             =   7920
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   249
      Left            =   570
      TabIndex        =   557
      Top             =   7920
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   248
      Left            =   330
      TabIndex        =   556
      Top             =   7920
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   247
      Left            =   90
      TabIndex        =   555
      Top             =   7920
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   246
      Left            =   2970
      TabIndex        =   554
      Top             =   7680
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   245
      Left            =   2730
      TabIndex        =   553
      Top             =   7680
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   244
      Left            =   2490
      TabIndex        =   552
      Top             =   7680
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   243
      Left            =   2250
      TabIndex        =   551
      Top             =   7680
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   242
      Left            =   2010
      TabIndex        =   550
      Top             =   7680
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   241
      Left            =   1770
      TabIndex        =   549
      Top             =   7680
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   240
      Left            =   1530
      TabIndex        =   548
      Top             =   7680
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   239
      Left            =   1290
      TabIndex        =   547
      Top             =   7680
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   238
      Left            =   1050
      TabIndex        =   546
      Top             =   7680
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   237
      Left            =   810
      TabIndex        =   545
      Top             =   7680
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   236
      Left            =   570
      TabIndex        =   544
      Top             =   7680
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   235
      Left            =   330
      TabIndex        =   543
      Top             =   7680
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   234
      Left            =   90
      TabIndex        =   542
      Top             =   7680
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   233
      Left            =   2970
      TabIndex        =   541
      Top             =   7440
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   232
      Left            =   2730
      TabIndex        =   540
      Top             =   7440
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   231
      Left            =   2490
      TabIndex        =   539
      Top             =   7440
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   230
      Left            =   2250
      TabIndex        =   538
      Top             =   7440
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   229
      Left            =   2010
      TabIndex        =   537
      Top             =   7440
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   228
      Left            =   1770
      TabIndex        =   536
      Top             =   7440
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   227
      Left            =   1530
      TabIndex        =   535
      Top             =   7440
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   226
      Left            =   1290
      TabIndex        =   534
      Top             =   7440
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   225
      Left            =   1050
      TabIndex        =   533
      Top             =   7440
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   224
      Left            =   810
      TabIndex        =   532
      Top             =   7440
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   223
      Left            =   570
      TabIndex        =   531
      Top             =   7440
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   222
      Left            =   330
      TabIndex        =   530
      Top             =   7440
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   221
      Left            =   90
      TabIndex        =   529
      Top             =   7440
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   220
      Left            =   2970
      TabIndex        =   528
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   219
      Left            =   2730
      TabIndex        =   527
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   218
      Left            =   2490
      TabIndex        =   526
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   217
      Left            =   2250
      TabIndex        =   525
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   216
      Left            =   2010
      TabIndex        =   524
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   215
      Left            =   1770
      TabIndex        =   523
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   214
      Left            =   1530
      TabIndex        =   522
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   213
      Left            =   1290
      TabIndex        =   521
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   212
      Left            =   1050
      TabIndex        =   520
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   211
      Left            =   810
      TabIndex        =   519
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   210
      Left            =   570
      TabIndex        =   518
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   209
      Left            =   330
      TabIndex        =   517
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   208
      Left            =   90
      TabIndex        =   516
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   207
      Left            =   2970
      TabIndex        =   515
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   206
      Left            =   2730
      TabIndex        =   514
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   205
      Left            =   2490
      TabIndex        =   513
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   204
      Left            =   2250
      TabIndex        =   512
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   203
      Left            =   2010
      TabIndex        =   511
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   202
      Left            =   1770
      TabIndex        =   510
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   201
      Left            =   1530
      TabIndex        =   509
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   200
      Left            =   1290
      TabIndex        =   508
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   199
      Left            =   1050
      TabIndex        =   507
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   198
      Left            =   810
      TabIndex        =   506
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   197
      Left            =   570
      TabIndex        =   505
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   196
      Left            =   330
      TabIndex        =   504
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   195
      Left            =   90
      TabIndex        =   503
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   194
      Left            =   2970
      TabIndex        =   502
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   193
      Left            =   2730
      TabIndex        =   501
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   192
      Left            =   2490
      TabIndex        =   500
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   191
      Left            =   2250
      TabIndex        =   499
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   190
      Left            =   2010
      TabIndex        =   498
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   189
      Left            =   1770
      TabIndex        =   497
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   188
      Left            =   1530
      TabIndex        =   496
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   187
      Left            =   1290
      TabIndex        =   495
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   186
      Left            =   1050
      TabIndex        =   494
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   185
      Left            =   810
      TabIndex        =   493
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   184
      Left            =   570
      TabIndex        =   492
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   183
      Left            =   330
      TabIndex        =   491
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   182
      Left            =   90
      TabIndex        =   490
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   181
      Left            =   2970
      TabIndex        =   489
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   180
      Left            =   2730
      TabIndex        =   488
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   179
      Left            =   2490
      TabIndex        =   487
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   178
      Left            =   2250
      TabIndex        =   486
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   177
      Left            =   2010
      TabIndex        =   485
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   176
      Left            =   1770
      TabIndex        =   484
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   175
      Left            =   1530
      TabIndex        =   483
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   174
      Left            =   1290
      TabIndex        =   482
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   173
      Left            =   1050
      TabIndex        =   481
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   172
      Left            =   810
      TabIndex        =   480
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   171
      Left            =   570
      TabIndex        =   479
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   170
      Left            =   330
      TabIndex        =   478
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   169
      Left            =   90
      TabIndex        =   477
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   168
      Left            =   2970
      TabIndex        =   476
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   167
      Left            =   2730
      TabIndex        =   475
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   166
      Left            =   2490
      TabIndex        =   474
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   165
      Left            =   2250
      TabIndex        =   473
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   164
      Left            =   2010
      TabIndex        =   472
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   163
      Left            =   1770
      TabIndex        =   471
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   162
      Left            =   1530
      TabIndex        =   470
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   161
      Left            =   1290
      TabIndex        =   469
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   160
      Left            =   1050
      TabIndex        =   468
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   159
      Left            =   810
      TabIndex        =   467
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   158
      Left            =   570
      TabIndex        =   466
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   157
      Left            =   330
      TabIndex        =   465
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   156
      Left            =   90
      TabIndex        =   464
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   155
      Left            =   2970
      TabIndex        =   463
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   154
      Left            =   2730
      TabIndex        =   462
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   153
      Left            =   2490
      TabIndex        =   461
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   152
      Left            =   2250
      TabIndex        =   460
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   151
      Left            =   2010
      TabIndex        =   459
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   150
      Left            =   1770
      TabIndex        =   458
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   149
      Left            =   1530
      TabIndex        =   457
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   148
      Left            =   1290
      TabIndex        =   456
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   147
      Left            =   1050
      TabIndex        =   455
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   146
      Left            =   810
      TabIndex        =   454
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   145
      Left            =   570
      TabIndex        =   453
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   144
      Left            =   330
      TabIndex        =   452
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   143
      Left            =   90
      TabIndex        =   451
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   142
      Left            =   2970
      TabIndex        =   450
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   141
      Left            =   2730
      TabIndex        =   449
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   140
      Left            =   2490
      TabIndex        =   448
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   139
      Left            =   2250
      TabIndex        =   447
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   138
      Left            =   2010
      TabIndex        =   446
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   137
      Left            =   1770
      TabIndex        =   445
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   136
      Left            =   1530
      TabIndex        =   444
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   135
      Left            =   1290
      TabIndex        =   443
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   134
      Left            =   1050
      TabIndex        =   442
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   133
      Left            =   810
      TabIndex        =   441
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   132
      Left            =   570
      TabIndex        =   440
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   131
      Left            =   330
      TabIndex        =   439
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   130
      Left            =   90
      TabIndex        =   438
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   129
      Left            =   2970
      TabIndex        =   437
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   128
      Left            =   2730
      TabIndex        =   436
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   127
      Left            =   2490
      TabIndex        =   435
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   126
      Left            =   2250
      TabIndex        =   434
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   125
      Left            =   2010
      TabIndex        =   433
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   124
      Left            =   1770
      TabIndex        =   432
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   123
      Left            =   1530
      TabIndex        =   431
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   122
      Left            =   1290
      TabIndex        =   430
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   121
      Left            =   1050
      TabIndex        =   429
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   120
      Left            =   810
      TabIndex        =   428
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   119
      Left            =   570
      TabIndex        =   427
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   118
      Left            =   330
      TabIndex        =   426
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   117
      Left            =   90
      TabIndex        =   425
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   116
      Left            =   2970
      TabIndex        =   424
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   115
      Left            =   2730
      TabIndex        =   423
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   114
      Left            =   2490
      TabIndex        =   422
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   113
      Left            =   2250
      TabIndex        =   421
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   112
      Left            =   2010
      TabIndex        =   420
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   111
      Left            =   1770
      TabIndex        =   419
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   110
      Left            =   1530
      TabIndex        =   418
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   109
      Left            =   1290
      TabIndex        =   417
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   108
      Left            =   1050
      TabIndex        =   416
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   107
      Left            =   810
      TabIndex        =   415
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   106
      Left            =   570
      TabIndex        =   414
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   105
      Left            =   330
      TabIndex        =   413
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   104
      Left            =   90
      TabIndex        =   412
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   103
      Left            =   2970
      TabIndex        =   411
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   102
      Left            =   2730
      TabIndex        =   410
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   101
      Left            =   2490
      TabIndex        =   409
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   100
      Left            =   2250
      TabIndex        =   408
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   99
      Left            =   2010
      TabIndex        =   407
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   98
      Left            =   1770
      TabIndex        =   406
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   97
      Left            =   1530
      TabIndex        =   405
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   96
      Left            =   1290
      TabIndex        =   404
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   95
      Left            =   1050
      TabIndex        =   403
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   94
      Left            =   810
      TabIndex        =   402
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   93
      Left            =   570
      TabIndex        =   401
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   92
      Left            =   330
      TabIndex        =   400
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   91
      Left            =   90
      TabIndex        =   399
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   90
      Left            =   2970
      TabIndex        =   398
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   89
      Left            =   2730
      TabIndex        =   397
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   88
      Left            =   2490
      TabIndex        =   396
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   87
      Left            =   2250
      TabIndex        =   395
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   86
      Left            =   2010
      TabIndex        =   394
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   85
      Left            =   1770
      TabIndex        =   393
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   84
      Left            =   1530
      TabIndex        =   392
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   83
      Left            =   1290
      TabIndex        =   211
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   82
      Left            =   1050
      TabIndex        =   210
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   81
      Left            =   810
      TabIndex        =   209
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   80
      Left            =   570
      TabIndex        =   208
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   79
      Left            =   330
      TabIndex        =   207
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   78
      Left            =   90
      TabIndex        =   206
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   77
      Left            =   2970
      TabIndex        =   205
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   76
      Left            =   2730
      TabIndex        =   204
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   75
      Left            =   2490
      TabIndex        =   203
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   74
      Left            =   2250
      TabIndex        =   202
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   73
      Left            =   2010
      TabIndex        =   201
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   72
      Left            =   1770
      TabIndex        =   200
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   71
      Left            =   1530
      TabIndex        =   199
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   70
      Left            =   1290
      TabIndex        =   198
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   69
      Left            =   1050
      TabIndex        =   197
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   68
      Left            =   810
      TabIndex        =   196
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   67
      Left            =   570
      TabIndex        =   195
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   66
      Left            =   330
      TabIndex        =   194
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   65
      Left            =   90
      TabIndex        =   193
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   64
      Left            =   2970
      TabIndex        =   192
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   63
      Left            =   2730
      TabIndex        =   191
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   62
      Left            =   2490
      TabIndex        =   190
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   61
      Left            =   2250
      TabIndex        =   189
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   60
      Left            =   2010
      TabIndex        =   188
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   59
      Left            =   1770
      TabIndex        =   91
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   58
      Left            =   1530
      TabIndex        =   90
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   57
      Left            =   1290
      TabIndex        =   89
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   56
      Left            =   1050
      TabIndex        =   88
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   55
      Left            =   810
      TabIndex        =   87
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   54
      Left            =   570
      TabIndex        =   86
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   53
      Left            =   330
      TabIndex        =   85
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   52
      Left            =   90
      TabIndex        =   84
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   51
      Left            =   2970
      TabIndex        =   83
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   50
      Left            =   2730
      TabIndex        =   82
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   49
      Left            =   2490
      TabIndex        =   81
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   48
      Left            =   2250
      TabIndex        =   80
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   47
      Left            =   2010
      TabIndex        =   79
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   46
      Left            =   1770
      TabIndex        =   78
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   45
      Left            =   1530
      TabIndex        =   77
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   44
      Left            =   1290
      TabIndex        =   76
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   43
      Left            =   1050
      TabIndex        =   75
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   42
      Left            =   810
      TabIndex        =   74
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   41
      Left            =   570
      TabIndex        =   73
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   40
      Left            =   330
      TabIndex        =   72
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   39
      Left            =   90
      TabIndex        =   71
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   38
      Left            =   2970
      TabIndex        =   70
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   37
      Left            =   2730
      TabIndex        =   69
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   36
      Left            =   2490
      TabIndex        =   68
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   35
      Left            =   2250
      TabIndex        =   67
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   34
      Left            =   2010
      TabIndex        =   66
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   33
      Left            =   1770
      TabIndex        =   65
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   32
      Left            =   1530
      TabIndex        =   64
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   31
      Left            =   1290
      TabIndex        =   63
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   30
      Left            =   1050
      TabIndex        =   62
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   29
      Left            =   810
      TabIndex        =   61
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   28
      Left            =   570
      TabIndex        =   60
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   27
      Left            =   330
      TabIndex        =   59
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   26
      Left            =   90
      TabIndex        =   58
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   25
      Left            =   2970
      TabIndex        =   57
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   24
      Left            =   2730
      TabIndex        =   56
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   23
      Left            =   2490
      TabIndex        =   55
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   22
      Left            =   2250
      TabIndex        =   54
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   21
      Left            =   2010
      TabIndex        =   53
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   20
      Left            =   1770
      TabIndex        =   52
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   19
      Left            =   1530
      TabIndex        =   51
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   18
      Left            =   1290
      TabIndex        =   50
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   17
      Left            =   1050
      TabIndex        =   49
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   16
      Left            =   810
      TabIndex        =   48
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   15
      Left            =   570
      TabIndex        =   47
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   14
      Left            =   330
      TabIndex        =   46
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   13
      Left            =   90
      TabIndex        =   45
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   12
      Left            =   2970
      TabIndex        =   44
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   11
      Left            =   2730
      TabIndex        =   43
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   10
      Left            =   2490
      TabIndex        =   42
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   9
      Left            =   2250
      TabIndex        =   41
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   8
      Left            =   2010
      TabIndex        =   40
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   7
      Left            =   1770
      TabIndex        =   39
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   1530
      TabIndex        =   38
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   1290
      TabIndex        =   37
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   1050
      TabIndex        =   36
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   3
      Left            =   810
      TabIndex        =   35
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   2
      Left            =   570
      TabIndex        =   34
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   330
      TabIndex        =   33
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   32
      Top             =   3360
      Width           =   255
   End
   Begin VB.Menu cmd_file 
      Caption         =   "&File"
      Begin VB.Menu cmd_mnu_open 
         Caption         =   "&open"
         Shortcut        =   ^O
      End
      Begin VB.Menu cmd_mnu_save 
         Caption         =   "&save"
         Shortcut        =   ^S
      End
      Begin VB.Menu cmd_mnu_print 
         Caption         =   "&print"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu cmd_about 
      Caption         =   "A&bout"
      NegotiatePosition=   2  'Middle
   End
   Begin VB.Menu cmd_exit 
      Caption         =   "exit"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i, lbl_num, h_m, h, Style, Ctxt, Help, Response As Integer
Dim marks_tur1 As Boolean
Dim Msg As Variant
Dim arr_chk1(0 To 24) As Integer
Dim Title, MyString As String
 
Private Sub cmd_chk1_Click(Index As Integer)
If cmd_chk1(Index).BackColor = &HFFFFFF Then
cmd_chk1(Index).BackColor = 0
Else
cmd_chk1(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_chk2_Click(Index As Integer)
If cmd_chk2(Index).BackColor = &HFFFFFF Then
cmd_chk2(Index).BackColor = 0
Else
cmd_chk2(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_chk3_Click(Index As Integer)
If cmd_chk3(Index).BackColor = &HFFFFFF Then
cmd_chk3(Index).BackColor = 0
Else
cmd_chk3(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_chk4_Click(Index As Integer)
If cmd_chk4(Index).BackColor = &HFFFFFF Then
cmd_chk4(Index).BackColor = 0
Else
cmd_chk4(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_chk5_Click(Index As Integer)
If cmd_chk5(Index).BackColor = &HFFFFFF Then
cmd_chk5(Index).BackColor = 0
Else
cmd_chk5(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_chk6_Click(Index As Integer)
If cmd_chk6(Index).BackColor = &HFFFFFF Then
cmd_chk6(Index).BackColor = 0
Else
cmd_chk6(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_chk7_Click(Index As Integer)
If cmd_chk7(Index).BackColor = &HFFFFFF Then
cmd_chk7(Index).BackColor = 0
Else
cmd_chk7(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_chk8_Click(Index As Integer)
If cmd_chk8(Index).BackColor = &HFFFFFF Then
cmd_chk8(Index).BackColor = 0
Else
cmd_chk8(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_chk9_Click(Index As Integer)
If cmd_chk9(Index).BackColor = &HFFFFFF Then
cmd_chk9(Index).BackColor = 0
Else
cmd_chk9(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_chk10_Click(Index As Integer)
If cmd_chk10(Index).BackColor = &HFFFFFF Then
cmd_chk10(Index).BackColor = 0
Else
cmd_chk10(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_chk11_Click(Index As Integer)
If cmd_chk11(Index).BackColor = &HFFFFFF Then
cmd_chk11(Index).BackColor = 0
Else
cmd_chk11(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_chk12_Click(Index As Integer)
If cmd_chk12(Index).BackColor = &HFFFFFF Then
cmd_chk12(Index).BackColor = 0
Else
cmd_chk12(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_chk13_Click(Index As Integer)
If cmd_chk13(Index).BackColor = &HFFFFFF Then
cmd_chk13(Index).BackColor = 0
Else
cmd_chk13(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_chk14_Click(Index As Integer)
If cmd_chk14(Index).BackColor = &HFFFFFF Then
cmd_chk14(Index).BackColor = 0
Else
cmd_chk14(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_chk15_Click(Index As Integer)
If cmd_chk15(Index).BackColor = &HFFFFFF Then
cmd_chk15(Index).BackColor = 0
Else
cmd_chk15(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_chk16_Click(Index As Integer)
If cmd_chk16(Index).BackColor = &HFFFFFF Then
cmd_chk16(Index).BackColor = 0
Else
cmd_chk16(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_chk17_Click(Index As Integer)
If cmd_chk17(Index).BackColor = &HFFFFFF Then
cmd_chk17(Index).BackColor = 0
Else
cmd_chk17(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_chk18_Click(Index As Integer)
If cmd_chk18(Index).BackColor = &HFFFFFF Then
cmd_chk18(Index).BackColor = 0
Else
cmd_chk18(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_chk19_Click(Index As Integer)
If cmd_chk19(Index).BackColor = &HFFFFFF Then
cmd_chk19(Index).BackColor = 0
Else
cmd_chk19(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_chk20_Click(Index As Integer)
If cmd_chk20(Index).BackColor = &HFFFFFF Then
cmd_chk20(Index).BackColor = 0
Else
cmd_chk20(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_chk21_Click(Index As Integer)
If cmd_chk21(Index).BackColor = &HFFFFFF Then
cmd_chk21(Index).BackColor = 0
Else
cmd_chk21(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_chk22_Click(Index As Integer)
If cmd_chk22(Index).BackColor = &HFFFFFF Then
cmd_chk22(Index).BackColor = 0
Else
cmd_chk22(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_chk23_Click(Index As Integer)
If cmd_chk23(Index).BackColor = &HFFFFFF Then
cmd_chk23(Index).BackColor = 0
Else
cmd_chk23(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_chk24_Click(Index As Integer)
If cmd_chk24(Index).BackColor = &HFFFFFF Then
cmd_chk24(Index).BackColor = 0
Else
cmd_chk24(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_chk25_Click(Index As Integer)
If cmd_chk25(Index).BackColor = &HFFFFFF Then
cmd_chk25(Index).BackColor = 0
Else
cmd_chk25(Index).BackColor = &HFFFFFF
End If
End Sub
Private Sub cmd_mnu_open_Click()
Call cmdOpen_Click
End Sub

Private Sub cmd_mnu_print_Click()
Call cmd_print_Click
End Sub

Private Sub cmd_mnu_save_Click()
Call cmdSave_Click
End Sub

Private Sub cmd_N_Click()

For i = 0 To 24
    If cmd_chk1(i).BackColor = 0 Then
     cmd_chk1(i).BackColor = &HFFFFFF
     Else
    cmd_chk1(i).BackColor = 0
    End If
Next

For i = 0 To 24
    If cmd_chk2(i).BackColor = 0 Then
     cmd_chk2(i).BackColor = &HFFFFFF
     Else
    cmd_chk2(i).BackColor = 0
    End If
Next

For i = 0 To 24
    If cmd_chk3(i).BackColor = 0 Then
     cmd_chk3(i).BackColor = &HFFFFFF
     Else
    cmd_chk3(i).BackColor = 0
    End If
Next

For i = 0 To 24
    If cmd_chk4(i).BackColor = 0 Then
     cmd_chk4(i).BackColor = &HFFFFFF
     Else
    cmd_chk4(i).BackColor = 0
    End If
Next

For i = 0 To 24
    If cmd_chk5(i).BackColor = 0 Then
     cmd_chk5(i).BackColor = &HFFFFFF
     Else
    cmd_chk5(i).BackColor = 0
    End If
Next

For i = 0 To 24
    If cmd_chk6(i).BackColor = 0 Then
     cmd_chk6(i).BackColor = &HFFFFFF
     Else
    cmd_chk6(i).BackColor = 0
    End If
Next

For i = 0 To 24
    If cmd_chk7(i).BackColor = 0 Then
     cmd_chk7(i).BackColor = &HFFFFFF
     Else
    cmd_chk7(i).BackColor = 0
    End If
Next

For i = 0 To 24
    If cmd_chk9(i).BackColor = 0 Then
     cmd_chk9(i).BackColor = &HFFFFFF
     Else
    cmd_chk9(i).BackColor = 0
    End If
Next

For i = 0 To 24
    If cmd_chk8(i).BackColor = 0 Then
     cmd_chk8(i).BackColor = &HFFFFFF
     Else
    cmd_chk8(i).BackColor = 0
    End If
Next

For i = 0 To 24
    If cmd_chk10(i).BackColor = 0 Then
     cmd_chk10(i).BackColor = &HFFFFFF
     Else
    cmd_chk10(i).BackColor = 0
    End If
Next

For i = 0 To 24
    If cmd_chk11(i).BackColor = 0 Then
     cmd_chk11(i).BackColor = &HFFFFFF
     Else
    cmd_chk11(i).BackColor = 0
    End If
Next

For i = 0 To 24
    If cmd_chk12(i).BackColor = 0 Then
     cmd_chk12(i).BackColor = &HFFFFFF
     Else
    cmd_chk12(i).BackColor = 0
    End If
Next

For i = 0 To 24
    If cmd_chk13(i).BackColor = 0 Then
     cmd_chk13(i).BackColor = &HFFFFFF
     Else
    cmd_chk13(i).BackColor = 0
    End If
Next

For i = 0 To 24
    If cmd_chk14(i).BackColor = 0 Then
     cmd_chk14(i).BackColor = &HFFFFFF
     Else
    cmd_chk14(i).BackColor = 0
    End If
Next

For i = 0 To 24
    If cmd_chk15(i).BackColor = 0 Then
     cmd_chk15(i).BackColor = &HFFFFFF
     Else
    cmd_chk15(i).BackColor = 0
    End If
Next

For i = 0 To 24
    If cmd_chk16(i).BackColor = 0 Then
     cmd_chk16(i).BackColor = &HFFFFFF
     Else
    cmd_chk16(i).BackColor = 0
    End If
Next

For i = 0 To 24
    If cmd_chk17(i).BackColor = 0 Then
     cmd_chk17(i).BackColor = &HFFFFFF
     Else
    cmd_chk17(i).BackColor = 0
    End If
Next

For i = 0 To 24
    If cmd_chk18(i).BackColor = 0 Then
     cmd_chk18(i).BackColor = &HFFFFFF
     Else
    cmd_chk18(i).BackColor = 0
    End If
Next

For i = 0 To 24
    If cmd_chk19(i).BackColor = 0 Then
     cmd_chk19(i).BackColor = &HFFFFFF
     Else
    cmd_chk19(i).BackColor = 0
    End If
Next

For i = 0 To 24
    If cmd_chk20(i).BackColor = 0 Then
     cmd_chk20(i).BackColor = &HFFFFFF
     Else
    cmd_chk20(i).BackColor = 0
    End If
Next

For i = 0 To 24
    If cmd_chk21(i).BackColor = 0 Then
     cmd_chk21(i).BackColor = &HFFFFFF
     Else
    cmd_chk21(i).BackColor = 0
    End If
Next

For i = 0 To 24
    If cmd_chk22(i).BackColor = 0 Then
     cmd_chk22(i).BackColor = &HFFFFFF
     Else
    cmd_chk22(i).BackColor = 0
    End If
Next

For i = 0 To 24
    If cmd_chk23(i).BackColor = 0 Then
     cmd_chk23(i).BackColor = &HFFFFFF
     Else
    cmd_chk23(i).BackColor = 0
    End If
Next

For i = 0 To 24
    If cmd_chk24(i).BackColor = 0 Then
     cmd_chk24(i).BackColor = &HFFFFFF
     Else
    cmd_chk24(i).BackColor = 0
    End If
Next

For i = 0 To 24
    If cmd_chk25(i).BackColor = 0 Then
     cmd_chk25(i).BackColor = &HFFFFFF
     Else
    cmd_chk25(i).BackColor = 0
    End If
Next
 Call cmd_calc_Click

End Sub


Private Sub cmd_rnd_Click()




If cmd_rnd.Caption = "&Stop !" Then

cmd_rnd.Caption = "&Random !"
tmr_rnd.Enabled = False
Call cmd_calc_Click
Else

Msg = "R U sure U want 2 do that ?"   ' Define message.
Style = vbYesNo + vbCritical + vbDefaultButton2 ' Define buttons.
Title = "Randoming All Of The Board ?"  ' Define title.
Ctxt = 1000 ' Define topic
        ' context.
        ' Display message.
Response = MsgBox(Msg, Style, Title, Help, Ctxt)

If Response = vbYes Then    ' User chose Yes.


tmr_rnd.Enabled = True
'Call tmr_rnd_Timer
cmd_rnd.Caption = "&Stop !"


For i = 0 To 12 'cmd_chk1
    Label1(i).Caption = ""
    Label2(i).Caption = ""
    Label3(i).Caption = ""
    Label4(i).Caption = ""
    Label5(i).Caption = ""
    Label6(i).Caption = ""
    Label7(i).Caption = ""
    Label8(i).Caption = ""
    Label9(i).Caption = ""
    Label10(i).Caption = ""
    Label11(i).Caption = ""
    Label12(i).Caption = ""
    Label13(i).Caption = ""
    Label14(i).Caption = ""
    Label15(i).Caption = ""
    Label16(i).Caption = ""
    Label17(i).Caption = ""
    Label18(i).Caption = ""
    Label19(i).Caption = ""
    Label20(i).Caption = ""
    Label21(i).Caption = ""
    Label22(i).Caption = ""
    Label23(i).Caption = ""
    Label24(i).Caption = ""
    Label25(i).Caption = ""
Next

For i = 0 To 324 'null all label26
    Label26(i).Caption = ""
Next
End If '(choosed yes)

End If '(else)

End Sub

Private Sub cmdOpen_Click() ' ##########################################



' המשתמש בחר לפתוח קובץ
Dim fileName As String
    
    CommonDlg.fileName = "" 'L אתחול שם הקובץ
    CommonDlg.Filter = "griddlers files (*.Gds) | *.Gds" 'L פילטר
    CommonDlg.Flags = cdlOFNHideReadOnly 'L הסתר את פתח לקריאה בלבד
    CommonDlg.ShowOpen 'L הצג את תיבת הדיאלוג פתיחה
    fileName = CommonDlg.fileName
    If fileName = "" Then Exit Sub 'L אם המשתמש בחר בביטול
    
    Text1.Text = "" ' איפוס תיבת הטקסט

    ' קרא לפרוצדורה לפתיחת קובץ לתיבת הטקסט
    
    Call OpenTextFile(cmd_chk1, fileName, Text1)
    
'מאפס את התאים
Call Unchecked_all(cmd_chk1, cmd_chk2, cmd_chk3, cmd_chk4, cmd_chk5, cmd_chk6, cmd_chk7, cmd_chk8, cmd_chk9, cmd_chk10, cmd_chk11, cmd_chk12, cmd_chk13, cmd_chk14, cmd_chk15, cmd_chk16, cmd_chk17, cmd_chk18, cmd_chk19, cmd_chk20, cmd_chk21, cmd_chk22, cmd_chk23, cmd_chk24, cmd_chk25, Label26)
' For h = 0 To 24

'cmd_chk1(h).BackColor  = &H00FFFFFF&
'cmd_chk2(h).BackColor  = &H00FFFFFF&
'cmd_chk3(h).BackColor  = &H00FFFFFF&
'cmd_chk4(h).BackColor  = &H00FFFFFF&
'cmd_chk5(h).BackColor  = &H00FFFFFF&
'cmd_chk6(h).BackColor  = &H00FFFFFF&
'cmd_chk7(h).BackColor  = &H00FFFFFF&
'cmd_chk8(h).BackColor  = &H00FFFFFF&
'cmd_chk9(h).BackColor  = &H00FFFFFF&
'cmd_chk10(h).BackColor  = &H00FFFFFF&
'cmd_chk11(h).BackColor  = &H00FFFFFF&
'cmd_chk12(h).BackColor  = &H00FFFFFF&
'cmd_chk13(h).BackColor  = &H00FFFFFF&
'cmd_chk14(h).BackColor  = &H00FFFFFF&
'cmd_chk15(h).BackColor  = &H00FFFFFF&
'cmd_chk16(h).BackColor  = &H00FFFFFF&
'cmd_chk17(h).BackColor  = &H00FFFFFF&
'cmd_chk18(h).BackColor  = &H00FFFFFF&
'cmd_chk19(h).BackColor  = &H00FFFFFF&
'cmd_chk20(h).BackColor  = &H00FFFFFF&
'cmd_chk21(h).BackColor  = &H00FFFFFF&
'cmd_chk22(h).BackColor  = &H00FFFFFF&
'cmd_chk23(h).BackColor  = &H00FFFFFF&
'cmd_chk24(h).BackColor  = &H00FFFFFF&
'cmd_chk25(h).BackColor  = &H00FFFFFF&

'Next
'
'For h = 0 To 324 ' null all label26
'Label26(h).Caption = ""
'Next
'מאפס את התאים סיום
    
    Dim last, next_line As Integer
        
        
        
next_line = 0
last = 1
i = 0
here1:
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
        If i = 25 Then
                GoTo next_line2
                'GoTo call_calc_here:
            End If
            cmd_chk1(i).BackColor = 0
            last = 1 + InStr(last, Text1.Text, 1) '!+i
            i = i + 1
            GoTo here1
        Else
            If i = 25 Then
                GoTo next_line2
                'GoTo call_calc_here:
            End If
            cmd_chk1(i).BackColor = &HFFFFFF
        End If
   Next
    
next_line2:
next_line = next_line + 25
last = 1 + 25 * 1
i = 0 + 25 * 1
here2:
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
                If i = next_line + 25 Then
                GoTo next_line3
                'GoTo call_calc_here:
            End If
            cmd_chk2(i - 25 * 1).BackColor = 0
            last = 1 + InStr(last, Text1.Text, 1) '!+i
            i = i + 1
            GoTo here2
        Else
            If i = 25 * 2 Then
                GoTo next_line3
                'GoTo call_calc_here:
            End If
            cmd_chk2(i - 25 * 1).BackColor = &HFFFFFF
        End If
   Next
    
     ' till  here  good!
    
next_line3:
next_line = next_line + 25
last = 1 + next_line
i = 0 + next_line
here3:
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
               If i = next_line + 25 Then
                GoTo next_line4
                'GoTo call_calc_here:
            End If
            cmd_chk3(i - next_line).BackColor = 0
            last = 1 + InStr(last, Text1.Text, 1) '!+i
            i = i + 1
            GoTo here3
        Else
            If i = next_line + 25 Then
                GoTo next_line4
                'GoTo call_calc_here:
            End If
            cmd_chk3(i - next_line).BackColor = &HFFFFFF
        End If
   Next
       
next_line4: 'change
next_line = next_line + 25
last = 1 + next_line
i = 0 + next_line
here4: 'change
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
                If i = next_line + 25 Then
                GoTo next_line5
                'GoTo call_calc_here:
            End If
            cmd_chk4(i - next_line).BackColor = 0    'change
            last = 1 + InStr(last, Text1.Text, 1) '!+i
            i = i + 1
            GoTo here4 'change
        Else
            If i = next_line + 25 Then
                GoTo next_line5 'change
                'GoTo call_calc_here:
            End If
            cmd_chk4(i - next_line).BackColor = &HFFFFFF    'change
        End If
   Next
          
next_line5: 'change
next_line = next_line + 25
last = 1 + next_line
i = 0 + next_line
here5: 'change
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
            If i = next_line + 25 Then
                GoTo next_line6  'change
                'GoTo call_calc_here:
            End If
            cmd_chk5(i - next_line).BackColor = 0     'change
            last = 1 + InStr(last, Text1.Text, 1) '!+i
            i = i + 1
            GoTo here5 'change
        Else
            If i = next_line + 25 Then
                GoTo next_line6 'change
                'GoTo call_calc_here:
            End If
            cmd_chk5(i - next_line).BackColor = &HFFFFFF     'change
        End If
   Next
   
next_line6: 'change
next_line = next_line + 25
last = 1 + next_line
i = 0 + next_line
here6: 'change
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
            If i = next_line + 25 Then
                GoTo next_line7  'change
                'GoTo call_calc_here:
            End If
            cmd_chk6(i - next_line).BackColor = 0    'change
            last = 1 + InStr(last, Text1.Text, 1) '!+i
            i = i + 1
            GoTo here6 'change
        Else
            If i = next_line + 25 Then
                GoTo next_line7 'change
                'GoTo call_calc_here:
            End If
            cmd_chk6(i - next_line).BackColor = &HFFFFFF  'change
        End If
   Next
    
 
next_line7: 'change
next_line = next_line + 25
last = 1 + next_line
i = 0 + next_line
here7: 'change
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
            If i = next_line + 25 Then
                GoTo next_line8  'change +1
                'GoTo call_calc_here:
            End If
            cmd_chk7(i - next_line).BackColor = 0     'change
            last = 1 + InStr(last, Text1.Text, 1) '!+i
            i = i + 1
            GoTo here7 'change
        Else
            If i = next_line + 25 Then
                GoTo next_line8 'change+1
                'GoTo call_calc_here:
            End If
            cmd_chk7(i - next_line).BackColor = &HFFFFFF  'change
        End If
   Next
    
next_line8: 'change
next_line = next_line + 25
last = 1 + next_line
i = 0 + next_line
here8: 'change
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
            If i = next_line + 25 Then
                GoTo next_line9  'change +1
                'GoTo call_calc_here:
            End If
            cmd_chk8(i - next_line).BackColor = 0     'change
            last = 1 + InStr(last, Text1.Text, 1) '!+i
            i = i + 1
            GoTo here8 'change
        Else
            If i = next_line + 25 Then
                GoTo next_line9 'change+1
                'GoTo call_calc_here:
            End If
            cmd_chk8(i - next_line).BackColor = &HFFFFFF  'change
        End If
   Next
  
next_line9: 'change
next_line = next_line + 25
last = 1 + next_line
i = 0 + next_line
here9: 'change
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
            If i = next_line + 25 Then
                GoTo next_line10  'change +1
                'GoTo call_calc_here:
            End If
            cmd_chk9(i - next_line).BackColor = 0     'change
            last = 1 + InStr(last, Text1.Text, 1) '!+i
            i = i + 1
            GoTo here9 'change
        Else
            If i = next_line + 25 Then
                GoTo next_line10 'change+1
                'GoTo call_calc_here:
            End If
            cmd_chk9(i - next_line).BackColor = &HFFFFFF  'change
        End If
   Next
   
next_line10:   'change
next_line = next_line + 25
last = 1 + next_line
i = 0 + next_line
here10: 'change
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
            If i = next_line + 25 Then
                GoTo next_line11 'change +1
                'GoTo call_calc_here:
            End If
            cmd_chk10(i - next_line).BackColor = 0     'change
            last = 1 + InStr(last, Text1.Text, 1) '!+i
            i = i + 1
            GoTo here10 'change
        Else
            If i = next_line + 25 Then
                GoTo next_line11 'change+1
                'GoTo call_calc_here:
            End If
            cmd_chk10(i - next_line).BackColor = &HFFFFFF  'change
        End If
   Next
   
next_line11: 'change
next_line = next_line + 25
last = 1 + next_line
i = 0 + next_line
here11: 'change
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
            If i = next_line + 25 Then
                GoTo next_line12  'change +1
                'GoTo call_calc_here:
            End If
            cmd_chk11(i - next_line).BackColor = 0     'change
            last = 1 + InStr(last, Text1.Text, 1) '!+i
            i = i + 1
            GoTo here11 'change
        Else
            If i = next_line + 25 Then
                GoTo next_line12 'change+1
                'GoTo call_calc_here:
            End If
            cmd_chk11(i - next_line).BackColor = &HFFFFFF  'change
        End If
   Next
   
next_line12:    'change
next_line = next_line + 25
last = 1 + next_line
i = 0 + next_line
here12: 'change
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
            If i = next_line + 25 Then
                GoTo next_line13 'change +1
                'GoTo call_calc_here:
            End If
            cmd_chk12(i - next_line).BackColor = 0     'change
            last = 1 + InStr(last, Text1.Text, 1) '!+i
            i = i + 1
            GoTo here12 'change
        Else
            If i = next_line + 25 Then
                GoTo next_line13 'change+1
                'GoTo call_calc_here:
            End If
            cmd_chk12(i - next_line).BackColor = &HFFFFFF  'change
        End If
   Next
       
next_line13:    'change
next_line = next_line + 25
last = 1 + next_line
i = 0 + next_line
here13: 'change
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
            If i = next_line + 25 Then
                GoTo next_line14 'change +1
                'GoTo call_calc_here:
            End If
            cmd_chk13(i - next_line).BackColor = 0     'change
            last = 1 + InStr(last, Text1.Text, 1) '!+i
            i = i + 1
            GoTo here13 'change
        Else
            If i = next_line + 25 Then
                GoTo next_line14 'change+1
                'GoTo call_calc_here:
            End If
            cmd_chk13(i - next_line).BackColor = &HFFFFFF  'change
        End If
        
   Next
        
next_line14:    'change
next_line = next_line + 25
last = 1 + next_line
i = 0 + next_line
here14: 'change
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
            If i = next_line + 25 Then
                GoTo next_line15 'change +1
                'GoTo call_calc_here:
            End If
            cmd_chk14(i - next_line).BackColor = 0     'change
            last = 1 + InStr(last, Text1.Text, 1) '!+i
            i = i + 1
            GoTo here14 'change
        Else
            If i = next_line + 25 Then
                GoTo next_line15 'change+1
                'GoTo call_calc_here:
            End If
            cmd_chk14(i - next_line).BackColor = &HFFFFFF  'change
        End If
   Next
           
next_line15:    'change
next_line = next_line + 25
last = 1 + next_line
i = 0 + next_line
here15: 'change
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
            If i = next_line + 25 Then
                GoTo next_line16 'change +1
                'GoTo call_calc_here:
            End If
            cmd_chk15(i - next_line).BackColor = 0     'change
            last = 1 + InStr(last, Text1.Text, 1) '!+i
            i = i + 1
            GoTo here15 'change
        Else
            If i = next_line + 25 Then
                GoTo next_line16 'change+1
                'GoTo call_calc_here:
            End If
            cmd_chk15(i - next_line).BackColor = &HFFFFFF  'change
        End If
   Next
   
           
next_line16:    'change
next_line = next_line + 25
last = 1 + next_line
i = 0 + next_line
here16: 'change
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
            If i = next_line + 25 Then
                GoTo next_line17 'change +1
                'GoTo call_calc_here:
            End If
            cmd_chk16(i - next_line).BackColor = 0     'change
            last = 1 + InStr(last, Text1.Text, 1) '!+i
            i = i + 1
            GoTo here16 'change
        Else
            If i = next_line + 25 Then
                GoTo next_line17 'change+1
                'GoTo call_calc_here:
            End If
            cmd_chk16(i - next_line).BackColor = &HFFFFFF  'change
        End If
   Next
            
next_line17:    'change
next_line = next_line + 25
last = 1 + next_line
i = 0 + next_line
here17: 'change
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
            If i = next_line + 25 Then
                GoTo next_line18 'change +1
                'GoTo call_calc_here:
            End If
            cmd_chk17(i - next_line).BackColor = 0     'change
            last = 1 + InStr(last, Text1.Text, 1) '!+i
            i = i + 1
            GoTo here17 'change
        Else
            If i = next_line + 25 Then
                GoTo next_line18 'change+1
                'GoTo call_calc_here:
            End If
            cmd_chk17(i - next_line).BackColor = &HFFFFFF  'change
        End If
   Next
            
next_line18:    'change
next_line = next_line + 25
last = 1 + next_line
i = 0 + next_line
here18: 'change
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
            If i = next_line + 25 Then
                GoTo next_line19 'change +1
                'GoTo call_calc_here:
            End If
            cmd_chk18(i - next_line).BackColor = 0     'change
            last = 1 + InStr(last, Text1.Text, 1) '!+i
            i = i + 1
            GoTo here18 'change
        Else
            If i = next_line + 25 Then
                GoTo next_line19 'change+1
                'GoTo call_calc_here:
            End If
            cmd_chk18(i - next_line).BackColor = &HFFFFFF  'change
        End If
   Next
            
next_line19:    'change
next_line = next_line + 25
last = 1 + next_line
i = 0 + next_line
here19: 'change
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
            If i = next_line + 25 Then
                GoTo next_line20 'change +1
                'GoTo call_calc_here:
            End If
            cmd_chk19(i - next_line).BackColor = 0     'change
            last = 1 + InStr(last, Text1.Text, 1) '!+i
            i = i + 1
            GoTo here19 'change
        Else
            If i = next_line + 25 Then
                GoTo next_line20 'change+1
                'GoTo call_calc_here:
            End If
            cmd_chk19(i - next_line).BackColor = &HFFFFFF  'change
        End If
   Next
 
            
next_line20:    'change
next_line = next_line + 25
last = 1 + next_line
i = 0 + next_line
here20: 'change
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
            If i = next_line + 25 Then
                GoTo next_line21 'change +1
                'GoTo call_calc_here:
            End If
            cmd_chk20(i - next_line).BackColor = 0     'change
            last = 1 + InStr(last, Text1.Text, 1) '!+i
            i = i + 1
            GoTo here20 'change
        Else
            If i = next_line + 25 Then
                GoTo next_line21 'change+1
                'GoTo call_calc_here:
            End If
            cmd_chk20(i - next_line).BackColor = &HFFFFFF  'change
        End If
   Next
             
next_line21:    'change
next_line = next_line + 25
last = 1 + next_line
i = 0 + next_line
here21: 'change
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
            If i = next_line + 25 Then
                GoTo next_line22 'change +1
                'GoTo call_calc_here:
            End If
            cmd_chk21(i - next_line).BackColor = 0     'change
            last = 1 + InStr(last, Text1.Text, 1)
            i = i + 1
            GoTo here21 'change
        Else
            If i = next_line + 25 Then
                GoTo next_line22 'change+1
                'GoTo call_calc_here:
            End If
            cmd_chk21(i - next_line).BackColor = &HFFFFFF  'change
        End If
   Next
   
                
next_line22:    'change
next_line = next_line + 25
last = 1 + next_line
i = 0 + next_line
here22: 'change
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
            If i = next_line + 25 Then
                GoTo next_line23 'change +1
                'GoTo call_calc_here:
            End If
            cmd_chk22(i - next_line).BackColor = 0     'change
            last = 1 + InStr(last, Text1.Text, 1)
            i = i + 1
            GoTo here22 'change
        Else
            If i = next_line + 25 Then
                GoTo next_line23 'change+1
                'GoTo call_calc_here:
            End If
            cmd_chk22(i - next_line).BackColor = &HFFFFFF  'change
        End If
   Next
                
next_line23:    'change
next_line = next_line + 25
last = 1 + next_line
i = 0 + next_line
here23: 'change
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
            If i = next_line + 25 Then
                GoTo next_line24 'change +1
                'GoTo call_calc_here:
            End If
            cmd_chk23(i - next_line).BackColor = 0     'change
            last = 1 + InStr(last, Text1.Text, 1)
            i = i + 1
            GoTo here23 'change
        Else
            If i = next_line + 25 Then
                GoTo next_line24 'change+1
                'GoTo call_calc_here:
            End If
            cmd_chk23(i - next_line).BackColor = &HFFFFFF  'change
        End If
   Next
                
next_line24:    'change
next_line = next_line + 25
last = 1 + next_line
i = 0 + next_line
here24: 'change
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
            If i = next_line + 25 Then
                GoTo next_line25 'change +1
                'GoTo call_calc_here:
            End If
            cmd_chk24(i - next_line).BackColor = 0     'change
            last = 1 + InStr(last, Text1.Text, 1)
            i = i + 1
            GoTo here24 'change
        Else
            If i = next_line + 25 Then
                GoTo next_line25 'change+1
                'GoTo call_calc_here:
            End If
            cmd_chk24(i - next_line).BackColor = &HFFFFFF  'change
        End If
   Next
                
next_line25:    'change
next_line = next_line + 25
last = 1 + next_line
i = 0 + next_line
here25: 'change
For i = i To InStr(last, Text1.Text, 1)
        If InStr(last, Text1.Text, 1) = i + 1 Then
            If i = next_line + 25 Then
                'GoTo next_line22 'change +1
                GoTo call_calc_here
            End If
            cmd_chk25(i - next_line).BackColor = 0     'change
            last = 1 + InStr(last, Text1.Text, 1)
            i = i + 1
            GoTo here25 'change
        Else
            If i = next_line + 25 Then
                'GoTo next_line22 'change+1
                GoTo call_calc_here:
            End If
            cmd_chk25(i - next_line).BackColor = &HFFFFFF  'change
        End If
   Next
   

call_calc_here:
Call cmd_calc_Click

     ' בדיקה האם אנו מחפשים טקסט זה פעם ראשונה או לא
 '   If TextToFind = sTarget Then
  '      TargetPos = TargetPos + 1
  '  Else
   '     TargetPos = 1
   ' End If
   ' sTarget = TextToFind
        ' מציאת מיקומו של הטקסט לחיפוש בתיבת הטקסט
   '     TargetPos = InStr(TargetPos , where, sTarget, 1)
 
   
    
End Sub

Private Sub cmdSave_Click()
' המשתמש בחר לשמור את תוכנה של תיבת הטקסט לקובץ
Dim Reply As Integer
Dim fileName As String

    CommonDlg.fileName = "" 'L אתחול שם הקובץ
   ' CommonDlg.Filter = "Text files (*.txt) | *.txt" 'L פילטר
    CommonDlg.Filter = "griddlers files (*.Gds) | *.Gds" 'L פילטר
    CommonDlg.Flags = cdlOFNHideReadOnly 'L הסתר את שמור לקריאה בלבד
    CommonDlg.ShowSave 'L הצג את תיבת הדיאלוג שמירה
    fileName = CommonDlg.fileName
    If fileName = "" Then Exit Sub 'L אם המשתמש בחר בביטול
    ' בדיקה האם הקובץ כבר קיים
    On Error Resume Next
    Name CommonDlg.fileName As CommonDlg.fileName
    If Err = 0 Then
        ' אם הקובץ קיים הצג הודעה האם להחליף אותו
        Reply = MsgBox("? Already exist a file with this name - to replace", vbMsgBoxRtlReading + vbMsgBoxRight + vbYesNo, "? Replacing file")
        If Reply = 7 Then Call cmdSave_Click  'L המשתמש בחר לא להחליף את הקובץ
    End If
    ' קריאה לפרוצדורה לשמירת תוכנה של תיבת הטקסט לקובץ
    Call SaveTextFile(cmd_chk1, cmd_chk2, cmd_chk3, cmd_chk4, cmd_chk5, cmd_chk6, cmd_chk7, cmd_chk8, cmd_chk9, cmd_chk10, cmd_chk11, cmd_chk12, cmd_chk13, cmd_chk14, cmd_chk15, cmd_chk16, cmd_chk17, cmd_chk18, cmd_chk19, cmd_chk20, cmd_chk21, cmd_chk22, cmd_chk23, cmd_chk24, cmd_chk25, fileName)
End Sub  ' ##########################################


''Private Sub cmd_chk1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'cmd_chk1(Index).BackColor = &HFFFF&
'End Sub


Private Sub cmd_about_Click()
Dim s As String

s = "This Sudanic program dedicated to Noam !" + vbCrLf + vbCrLf + "connect eliran_t@walla.com"
MsgBox s
End Sub

Private Sub cmd_print_Click() ' ############################ printing


Frame1.Visible = True
Form1.BackColor = &H80000005
Frame1.BackColor = &H80000005
 
 For i = 0 To 2
Frame2(i).Visible = True
Next

For i = 0 To 12 'cmd_chk1
    
    Label1(i).BackColor = &HFFFFFF
    Label2(i).BackColor = &HFFFFFF
    Label3(i).BackColor = &HFFFFFF
    Label4(i).BackColor = &HFFFFFF
    Label5(i).BackColor = &HFFFFFF
    Label6(i).BackColor = &HFFFFFF
    Label7(i).BackColor = &HFFFFFF
    Label8(i).BackColor = &HFFFFFF
    Label9(i).BackColor = &HFFFFFF
    Label10(i).BackColor = &HFFFFFF
    Label11(i).BackColor = &HFFFFFF
    Label12(i).BackColor = &HFFFFFF
    Label13(i).BackColor = &HFFFFFF
    Label14(i).BackColor = &HFFFFFF
    Label15(i).BackColor = &HFFFFFF
    Label16(i).BackColor = &HFFFFFF
    Label17(i).BackColor = &HFFFFFF
    Label18(i).BackColor = &HFFFFFF
    Label19(i).BackColor = &HFFFFFF
    Label20(i).BackColor = &HFFFFFF
    Label21(i).BackColor = &HFFFFFF
    Label22(i).BackColor = &HFFFFFF
    Label23(i).BackColor = &HFFFFFF
    Label24(i).BackColor = &HFFFFFF
    Label25(i).BackColor = &HFFFFFF
Next

For h = 0 To 324 ' null all label26
Label26(h).BackColor = &HFFFFFF
Next

For i = 0 To 24
    
    cmd_chk1(i).Visible = False
    cmd_chk2(i).Visible = False
    cmd_chk3(i).Visible = False
    cmd_chk4(i).Visible = False
    cmd_chk5(i).Visible = False
    cmd_chk6(i).Visible = False
    cmd_chk7(i).Visible = False
    cmd_chk8(i).Visible = False
    cmd_chk9(i).Visible = False
    cmd_chk10(i).Visible = False
    cmd_chk11(i).Visible = False
    cmd_chk12(i).Visible = False
    cmd_chk13(i).Visible = False
    cmd_chk14(i).Visible = False
    cmd_chk15(i).Visible = False
    cmd_chk16(i).Visible = False
    cmd_chk17(i).Visible = False
    cmd_chk18(i).Visible = False
    cmd_chk19(i).Visible = False
    cmd_chk20(i).Visible = False
    cmd_chk21(i).Visible = False
    cmd_chk22(i).Visible = False
    cmd_chk23(i).Visible = False
    cmd_chk24(i).Visible = False
    cmd_chk25(i).Visible = False
Next

'On Error GoTo no_print
'Form1.PrintForm '@@@@@@@@@@@@@@@ ההדפסה עצמה

'no_print:
'If Error = "Printer error" Then
'MsgBox ("you have no printer !")
'End If

On Error GoTo no_print
Form1.PrintForm '@@@@@@@@@@@@@@@ ההדפסה עצמה

no_print:
If Error = "Printer error" Then
MsgBox ("you have no printer !")
End If


Frame1.Visible = False
Frame1.BackColor = &H80000013
Form1.BackColor = &HC000&

 For i = 0 To 2
Frame2(i).Visible = False
Next

For i = 0 To 24
    
    cmd_chk1(i).Visible = True
    cmd_chk2(i).Visible = True
    cmd_chk3(i).Visible = True
    cmd_chk4(i).Visible = True
    cmd_chk5(i).Visible = True
    cmd_chk6(i).Visible = True
    cmd_chk7(i).Visible = True
    cmd_chk8(i).Visible = True
    cmd_chk9(i).Visible = True
    cmd_chk10(i).Visible = True
    cmd_chk11(i).Visible = True
    cmd_chk12(i).Visible = True
    cmd_chk13(i).Visible = True
    cmd_chk14(i).Visible = True
    cmd_chk15(i).Visible = True
    cmd_chk16(i).Visible = True
    cmd_chk17(i).Visible = True
    cmd_chk18(i).Visible = True
    cmd_chk19(i).Visible = True
    cmd_chk20(i).Visible = True
    cmd_chk21(i).Visible = True
    cmd_chk22(i).Visible = True
    cmd_chk23(i).Visible = True
    cmd_chk24(i).Visible = True
    cmd_chk25(i).Visible = True
Next

For i = 0 To 12 'cmd_chk1
    Label1(i).BackColor = &HC0FFC0
    Label2(i).BackColor = &HC0FFC0
    Label3(i).BackColor = &HC0FFC0
    Label4(i).BackColor = &HC0FFC0
    Label5(i).BackColor = &HC0FFC0
    Label6(i).BackColor = &HC0FFC0
    Label7(i).BackColor = &HC0FFC0
    Label8(i).BackColor = &HC0FFC0
    Label9(i).BackColor = &HC0FFC0
    Label10(i).BackColor = &HC0FFC0
    Label11(i).BackColor = &HC0FFC0
    Label12(i).BackColor = &HC0FFC0
    Label13(i).BackColor = &HC0FFC0
    Label14(i).BackColor = &HC0FFC0
    Label15(i).BackColor = &HC0FFC0
    Label16(i).BackColor = &HC0FFC0
    Label17(i).BackColor = &HC0FFC0
    Label18(i).BackColor = &HC0FFC0
    Label19(i).BackColor = &HC0FFC0
    Label20(i).BackColor = &HC0FFC0
    Label21(i).BackColor = &HC0FFC0
    Label22(i).BackColor = &HC0FFC0
    Label23(i).BackColor = &HC0FFC0
    Label24(i).BackColor = &HC0FFC0
    Label25(i).BackColor = &HC0FFC0
Next

For h = 0 To 324 ' null all label26
Label26(h).BackColor = &HC0FFC0
Next



End Sub

' ############################# any typong on cmd_chkbox call cmd_calc_Click



'Private Sub cmd_chk1_Click(Index As Integer)
'call cmd_calc_Click

'End Sub
'Private Sub cmd_chk2_Click(Index As Integer)
'call cmd_calc_Click

'End Sub
'Private Sub cmd_chk3_Click(Index As Integer)
'call cmd_calc_Click

'End Sub
'Private Sub cmd_chk4_Click(Index As Integer)
'call cmd_calc_Click

'End Sub
'Private Sub cmd_chk5_Click(Index As Integer)
'call cmd_calc_Click

'End Sub
'Private Sub cmd_chk6_Click(Index As Integer)
'call cmd_calc_Click

'End Sub
'Private Sub cmd_chk7_Click(Index As Integer)
'call cmd_calc_Click

'End Sub
'Private Sub cmd_chk8_Click(Index As Integer)
'call cmd_calc_Click

'End Sub
'Private Sub cmd_chk9_Click(Index As Integer)
'call cmd_calc_Click

'End Sub
'Private Sub cmd_chk10_Click(Index As Integer)
'call cmd_calc_Click

'End Sub
'Private Sub cmd_chk11_Click(Index As Integer)
'call cmd_calc_Click

'End Sub
'Private Sub cmd_chk12_Click(Index As Integer)
'call cmd_calc_Click

'End Sub
'Private Sub cmd_chk13_Click(Index As Integer)
'call cmd_calc_Click

'End Sub
'Private Sub cmd_chk14_Click(Index As Integer)
'call cmd_calc_Click

'End Sub
'Private Sub cmd_chk15_Click(Index As Integer)
'call cmd_calc_Click

'End Sub
'Private Sub cmd_chk16_Click(Index As Integer)
'call cmd_calc_Click

'End Sub
'Private Sub cmd_chk17_Click(Index As Integer)
'call cmd_calc_Click

'End Sub
'Private Sub cmd_chk18_Click(Index As Integer)
'call cmd_calc_Click

'End Sub
'Private Sub cmd_chk19_Click(Index As Integer)
'call cmd_calc_Click

'End Sub
'Private Sub cmd_chk20_Click(Index As Integer)
'call cmd_calc_Click

'End Sub
'Private Sub cmd_chk21_Click(Index As Integer)
'call cmd_calc_Click

'End Sub
'Private Sub cmd_chk22_Click(Index As Integer)
'call cmd_calc_Click

'End Sub
'Private Sub cmd_chk23_Click(Index As Integer)
'call cmd_calc_Click

'End Sub
'Private Sub cmd_chk24_Click(Index As Integer)
'call cmd_calc_Click

'End Sub
'Private Sub cmd_chk25_Click(Index As Integer)
'call cmd_calc_Click

'End Sub

' ############################# any typong on cmd_chkbox call cmd_calc_Click

Private Sub cmd_cln_Click()

If tmr_rnd.Enabled = True Then
tmr_rnd.Enabled = False
cmd_rnd.Caption = "&Random !"
End If

Msg = "R U sure U want 2 do that ?"   ' Define message.
Style = vbYesNo + vbCritical + vbDefaultButton2 ' Define buttons.
Title = "Cleaning Board ?"  ' Define title.
Ctxt = 1000 ' Define topic
        ' context.
        ' Display message.
Response = MsgBox(Msg, Style, Title, Help, Ctxt)

If Response = vbYes Then    ' User chose Yes.

For i = 0 To 12 'cmd_chk1
    Label1(i).Caption = ""
    Label2(i).Caption = ""
    Label3(i).Caption = ""
    Label4(i).Caption = ""
    Label5(i).Caption = ""
    Label6(i).Caption = ""
    Label7(i).Caption = ""
    Label8(i).Caption = ""
    Label9(i).Caption = ""
    Label10(i).Caption = ""
    Label11(i).Caption = ""
    Label12(i).Caption = ""
    Label13(i).Caption = ""
    Label14(i).Caption = ""
    Label15(i).Caption = ""
    Label16(i).Caption = ""
    Label17(i).Caption = ""
    Label18(i).Caption = ""
    Label19(i).Caption = ""
    Label20(i).Caption = ""
    Label21(i).Caption = ""
    Label22(i).Caption = ""
    Label23(i).Caption = ""
    Label24(i).Caption = ""
    Label25(i).Caption = ""
Next

For h = 0 To 324 ' null all label26
Label26(h).Caption = ""
Next


For h = 0 To 24

cmd_chk1(h).BackColor = &HFFFFFF
cmd_chk2(h).BackColor = &HFFFFFF
cmd_chk3(h).BackColor = &HFFFFFF
cmd_chk4(h).BackColor = &HFFFFFF
cmd_chk5(h).BackColor = &HFFFFFF
cmd_chk6(h).BackColor = &HFFFFFF
cmd_chk7(h).BackColor = &HFFFFFF
cmd_chk8(h).BackColor = &HFFFFFF
cmd_chk9(h).BackColor = &HFFFFFF
cmd_chk10(h).BackColor = &HFFFFFF
cmd_chk11(h).BackColor = &HFFFFFF
cmd_chk12(h).BackColor = &HFFFFFF
cmd_chk13(h).BackColor = &HFFFFFF
cmd_chk14(h).BackColor = &HFFFFFF
cmd_chk15(h).BackColor = &HFFFFFF
cmd_chk16(h).BackColor = &HFFFFFF
cmd_chk17(h).BackColor = &HFFFFFF
cmd_chk18(h).BackColor = &HFFFFFF
cmd_chk19(h).BackColor = &HFFFFFF
cmd_chk20(h).BackColor = &HFFFFFF
cmd_chk21(h).BackColor = &HFFFFFF
cmd_chk22(h).BackColor = &HFFFFFF
cmd_chk23(h).BackColor = &HFFFFFF
cmd_chk24(h).BackColor = &HFFFFFF
cmd_chk25(h).BackColor = &HFFFFFF

Next



Else    ' User chose No.
    MyString = "No" ' Perform some action.
End If



End Sub


'Private Sub cmd_exit_Click()
'Msg = "R U sure U want 2 do that ?"   ' Define message.
'Style = vbYesNo + vbCritical + vbDefaultButton2 ' Define buttons.
'Title = "Mark All Board ?"  ' Define title.
'Ctxt = 1000 ' Define topic
        ' context.
        ' Display message.
'Response = MsgBox(Msg, Style, Title, Help, Ctxt)

'If Response = vbYes Then    ' User chose Yes.
'End If

'End Sub

Private Sub cmd_mark_Click()
' "R U sure U want 2 do that ?"

If tmr_rnd.Enabled = True Then
tmr_rnd.Enabled = False
cmd_rnd.Caption = "&Random !"
End If

Msg = "R U sure U want 2 do that ?"   ' Define message.
Style = vbYesNo + vbCritical + vbDefaultButton2 ' Define buttons.
Title = "Mark All Board ?"  ' Define title.
Ctxt = 1000 ' Define topic
        ' context.
        ' Display message.
Response = MsgBox(Msg, Style, Title, Help, Ctxt)

If Response = vbYes Then    ' User chose Yes.

For i = 0 To 12 'cmd_chk1
    Label1(i).Caption = ""
    Label2(i).Caption = ""
    Label3(i).Caption = ""
    Label4(i).Caption = ""
    Label5(i).Caption = ""
    Label6(i).Caption = ""
    Label7(i).Caption = ""
    Label8(i).Caption = ""
    Label9(i).Caption = ""
    Label10(i).Caption = ""
    Label11(i).Caption = ""
    Label12(i).Caption = ""
    Label13(i).Caption = ""
    Label14(i).Caption = ""
    Label15(i).Caption = ""
    Label16(i).Caption = ""
    Label17(i).Caption = ""
    Label18(i).Caption = ""
    Label19(i).Caption = ""
    Label20(i).Caption = ""
    Label21(i).Caption = ""
    Label22(i).Caption = ""
    Label23(i).Caption = ""
    Label24(i).Caption = ""
    Label25(i).Caption = ""
 
Next

For i = 0 To 324
    Label26(i).Caption = ""
Next


For h = 0 To 24
cmd_chk1(h).BackColor = 0
cmd_chk2(h).BackColor = 0
cmd_chk3(h).BackColor = 0
cmd_chk4(h).BackColor = 0
cmd_chk5(h).BackColor = 0
cmd_chk6(h).BackColor = 0
cmd_chk7(h).BackColor = 0
cmd_chk8(h).BackColor = 0
cmd_chk9(h).BackColor = 0
cmd_chk10(h).BackColor = 0
cmd_chk11(h).BackColor = 0
cmd_chk12(h).BackColor = 0
cmd_chk13(h).BackColor = 0
cmd_chk14(h).BackColor = 0
cmd_chk15(h).BackColor = 0
cmd_chk16(h).BackColor = 0
cmd_chk17(h).BackColor = 0
cmd_chk18(h).BackColor = 0
cmd_chk19(h).BackColor = 0
cmd_chk20(h).BackColor = 0
cmd_chk21(h).BackColor = 0
cmd_chk22(h).BackColor = 0
cmd_chk23(h).BackColor = 0
cmd_chk24(h).BackColor = 0
cmd_chk25(h).BackColor = 0
Next

End If

 Call cmd_calc_Click

End Sub


Private Sub cmd_calc_Click()
Dim does As Boolean
Dim does_line As Boolean
Dim i, lbl_num, h_m, h As Integer


'TURIM! #################################
 i = 0

 For i = 0 To 12 'cmd_chk1
    Label1(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk1(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur1
    End If
    Loop
    
sof_tur1:
    
    If does = True Then
    Label1(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next
    
 For i = 0 To 12 'cmd_chk2
    Label2(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk2(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur2
    End If
    Loop
    
sof_tur2:
    
    If does = True Then
    Label2(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next
    
    
 For i = 0 To 12 'cmd_chk3
    Label3(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk3(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur3
    End If
    Loop
    
sof_tur3:
    
    If does = True Then
    Label3(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next

 For i = 0 To 12 'cmd_chk4
    Label4(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk4(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur4
    End If
    Loop
    
sof_tur4:
    
    If does = True Then
    Label4(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next
    
    
For i = 0 To 12 'cmd_chk5
    Label5(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk5(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur5
    End If
    Loop
    
sof_tur5:
    
    If does = True Then
    Label5(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next

For i = 0 To 12 'cmd_chk6
    Label6(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk6(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur6
    End If
    Loop
    
sof_tur6:
    
    If does = True Then
    Label6(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next
    
For i = 0 To 12 'cmd_chk7
    Label7(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk7(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur7
    End If
    Loop
    
sof_tur7:
    
    If does = True Then
    Label7(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next
    
For i = 0 To 12 'cmd_chk8
    Label8(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk8(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur8
    End If
    Loop
    
sof_tur8:
    
    If does = True Then
    Label8(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next

For i = 0 To 12 'cmd_chk9
    Label9(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk9(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur9
    End If
    Loop
    
sof_tur9:
    
    If does = True Then
    Label9(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next

For i = 0 To 12 'cmd_chk10
    Label10(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk10(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur10
    End If
    Loop
    
sof_tur10:
    
    If does = True Then
    Label10(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next
    
For i = 0 To 12 'cmd_chk11
    Label11(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk11(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur11
    End If
    Loop
    
sof_tur11:
    
    If does = True Then
    Label11(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next
    
For i = 0 To 12 'cmd_chk12
    Label12(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk12(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur12
    End If
    Loop
    
sof_tur12:
    
    If does = True Then
    Label12(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next

For i = 0 To 12 'cmd_chk13
    Label13(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk13(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur13
    End If
    Loop
    
sof_tur13:
    
    If does = True Then
    Label13(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next

For i = 0 To 12 'cmd_chk14
    Label14(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk14(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur14
    End If
    Loop
    
sof_tur14:
    
    If does = True Then
    Label14(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next

For i = 0 To 12 'cmd_chk15
    Label15(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk15(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur15
    End If
    Loop
    
sof_tur15:
    
    If does = True Then
    Label15(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next

For i = 0 To 12 'cmd_chk16
    Label16(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk16(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur16
    End If
    Loop
    
sof_tur16:
    
    If does = True Then
    Label16(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next

For i = 0 To 12 'cmd_chk17
    Label17(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk17(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur17
    End If
    Loop
    
sof_tur17:
    
    If does = True Then
    Label17(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next

For i = 0 To 12 'cmd_chk18
    Label18(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk18(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur18
    End If
    Loop
    
sof_tur18:
    
    If does = True Then
    Label18(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next

For i = 0 To 12 'cmd_chk19
    Label19(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk19(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur19
    End If
    Loop
    
sof_tur19:
    
    If does = True Then
    Label19(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next

For i = 0 To 12 'cmd_chk20
    Label20(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk20(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur20
    End If
    Loop
    
sof_tur20:
    
    If does = True Then
    Label20(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next

For i = 0 To 12 'cmd_chk21
    Label21(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk21(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur21
    End If
    Loop
    
sof_tur21:
    
    If does = True Then
    Label21(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next

For i = 0 To 12 'cmd_chk22
    Label22(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk22(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur22
    End If
    Loop
    
sof_tur22:
    
    If does = True Then
    Label22(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next

For i = 0 To 12 'cmd_chk23
    Label23(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk23(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur23
    End If
    Loop
    
sof_tur23:
    
    If does = True Then
    Label23(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next

For i = 0 To 12 'cmd_chk24
    Label24(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk24(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur24
    End If
    Loop
    
sof_tur24:
    
    If does = True Then
    Label24(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next

For i = 0 To 12 'cmd_chk25
    Label25(i).Caption = ""
Next

 lbl_num = 0
 For i = 0 To 24
    
    does = False
    h_m = 0
 
    Do While cmd_chk25(i).BackColor = 0
    h_m = h_m + 1
    does = True
    i = i + 1
    If i = 25 Then
    GoTo sof_tur25
    End If
    Loop
    
sof_tur25:
    
    If does = True Then
    Label25(lbl_num).Caption = h_m
    lbl_num = lbl_num + 1
    i = i - 1
    End If
    
Next
'sof calc TURIM! ###################################################


'


'fixig labels TURIM! ###################################################

Dim h_f, j As Integer
Dim arr_fix(12, 1) As Integer

For i = 0 To 12
arr_fix(i, 0) = 0
arr_fix(i, 1) = 0
Next

h_f = 0  ' Label1

For i = 0 To 12
    If Not (Label1(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label1(i).Caption)
        Label1(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label1((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next

h_f = 0  ' Label2

For i = 0 To 12
    If Not (Label2(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label2(i).Caption)
        Label2(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label2((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next

h_f = 0  ' Label3

For i = 0 To 12
    If Not (Label3(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label3(i).Caption)
        Label3(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label3((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next

h_f = 0  ' Label4

For i = 0 To 12
    If Not (Label4(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label4(i).Caption)
        Label4(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label4((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next

h_f = 0  ' Label5

For i = 0 To 12
    If Not (Label5(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label5(i).Caption)
        Label5(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label5((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next

h_f = 0  ' Label6

For i = 0 To 12
    If Not (Label6(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label6(i).Caption)
        Label6(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label6((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next

h_f = 0  ' Label7

For i = 0 To 12
    If Not (Label7(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label7(i).Caption)
        Label7(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label7((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next

h_f = 0  ' Label8

For i = 0 To 12
    If Not (Label8(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label8(i).Caption)
        Label8(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label8((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next

h_f = 0  ' Label9

For i = 0 To 12
    If Not (Label9(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label9(i).Caption)
        Label9(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label9((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next

h_f = 0  ' Label10

For i = 0 To 12
    If Not (Label10(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label10(i).Caption)
        Label10(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label10((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next

h_f = 0  ' Label11

For i = 0 To 12
    If Not (Label11(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label11(i).Caption)
        Label11(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label11((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next

h_f = 0  ' Label12

For i = 0 To 12
    If Not (Label12(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label12(i).Caption)
        Label12(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label12((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next

h_f = 0  ' Label13

For i = 0 To 12
    If Not (Label13(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label13(i).Caption)
        Label13(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label13((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next

h_f = 0  ' Label14

For i = 0 To 12
    If Not (Label14(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label14(i).Caption)
        Label14(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label14((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next

h_f = 0  ' Label15

For i = 0 To 12
    If Not (Label15(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label15(i).Caption)
        Label15(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label15((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next

h_f = 0  ' Label16

For i = 0 To 12
    If Not (Label16(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label16(i).Caption)
        Label16(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label16((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next

h_f = 0  ' Label17

For i = 0 To 12
    If Not (Label17(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label17(i).Caption)
        Label17(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label17((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next

h_f = 0  ' Label18

For i = 0 To 12
    If Not (Label18(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label18(i).Caption)
        Label18(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label18((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next

h_f = 0  ' Label19

For i = 0 To 12
    If Not (Label19(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label19(i).Caption)
        Label19(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label19((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next

h_f = 0  ' Label20

For i = 0 To 12
    If Not (Label20(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label20(i).Caption)
        Label20(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label20((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next

h_f = 0  ' Label21

For i = 0 To 12
    If Not (Label21(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label21(i).Caption)
        Label21(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label21((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next

h_f = 0  ' Label222

For i = 0 To 12
    If Not (Label22(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label22(i).Caption)
        Label22(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label22((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next

h_f = 0  ' Label23

For i = 0 To 12
    If Not (Label23(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label23(i).Caption)
        Label23(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label23((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next

h_f = 0  ' Label24

For i = 0 To 12
    If Not (Label24(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label24(i).Caption)
        Label24(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label24((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next

h_f = 0  ' Label25

For i = 0 To 12
    If Not (Label25(i).Caption = "") Then
        arr_fix(i, 0) = i ' index
        arr_fix(i, 1) = CStr(Label25(i).Caption)
        Label25(i).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1
    Label25((12 - h_f + 1 + arr_fix(j, 0))).Caption = Str(arr_fix(j, 1))
'    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
Next







' sof fixig labels TURIM! ###################################################


'calc LINES! ###################################################


For i = 0 To 324 'null all label26
    Label26(i).Caption = ""
Next




lbl_num = 0


For h = 0 To 24 ' line

If h = 0 Then
lbl_num = h * 12
Else
lbl_num = (h * 12) + h
End If
h_m = 0
does_line = False

If cmd_chk1(h).BackColor = 0 Then
    h_m = h_m + 1
    Label26(lbl_num).Caption = h_m
    does_line = True
    GoTo next_cmd_chk2
Else
If does_line = True Then
    lbl_num = lbl_num + 1 'move next label
    h_m = 0
    does_line = False
End If
End If


next_cmd_chk2:
If cmd_chk2(h).BackColor = 0 Then
    h_m = h_m + 1
    Label26(lbl_num).Caption = h_m
    
    does_line = True
    GoTo next_cmd_chk3
Else

    If does_line = True Then
        lbl_num = lbl_num + 1 'move next label
        h_m = 0
        does_line = False
    End If

End If

next_cmd_chk3:
If cmd_chk3(h).BackColor = 0 Then
    h_m = h_m + 1
    Label26(lbl_num).Caption = h_m
    does_line = True
    GoTo next_cmd_chk4
Else
    If does_line = True Then
        lbl_num = lbl_num + 1 'move next label
        h_m = 0
        does_line = False
    End If
End If

next_cmd_chk4:
If cmd_chk4(h).BackColor = 0 Then
    h_m = h_m + 1
    Label26(lbl_num).Caption = h_m
    does_line = True
    GoTo next_cmd_chk5
Else
    If does_line = True Then
        lbl_num = lbl_num + 1 'move next label
        h_m = 0
        does_line = False
    End If
End If

next_cmd_chk5:
If cmd_chk5(h).BackColor = 0 Then
    h_m = h_m + 1
    Label26(lbl_num).Caption = h_m
    does_line = True
    GoTo next_cmd_chk6
Else
    If does_line = True Then
        lbl_num = lbl_num + 1 'move next label
        h_m = 0
        does_line = False
    End If
End If

next_cmd_chk6:
If cmd_chk6(h).BackColor = 0 Then
    h_m = h_m + 1
    Label26(lbl_num).Caption = h_m
    does_line = True
    GoTo next_cmd_chk7
Else
    If does_line = True Then
        lbl_num = lbl_num + 1 'move next label
        h_m = 0
        does_line = False
    End If
End If

next_cmd_chk7:
If cmd_chk7(h).BackColor = 0 Then
    h_m = h_m + 1
    Label26(lbl_num).Caption = h_m
    does_line = True
    does_line = True
    GoTo next_cmd_chk8
Else
    If does_line = True Then
        lbl_num = lbl_num + 1 'move next label
        h_m = 0
        does_line = False
    End If
End If

next_cmd_chk8:
If cmd_chk8(h).BackColor = 0 Then
    h_m = h_m + 1
    Label26(lbl_num).Caption = h_m
    does_line = True
    GoTo next_cmd_chk9
Else
    If does_line = True Then
        lbl_num = lbl_num + 1 'move next label
        h_m = 0
        does_line = False
    End If
End If

next_cmd_chk9:
If cmd_chk9(h).BackColor = 0 Then
    h_m = h_m + 1
    Label26(lbl_num).Caption = h_m
    does_line = True
    GoTo next_cmd_chk10
Else
    If does_line = True Then
        lbl_num = lbl_num + 1 'move next label
        h_m = 0
        does_line = False
    End If
End If

next_cmd_chk10:
If cmd_chk10(h).BackColor = 0 Then
    h_m = h_m + 1
    Label26(lbl_num).Caption = h_m
    does_line = True
    GoTo next_cmd_chk11
Else
    If does_line = True Then
        lbl_num = lbl_num + 1 'move next label
        h_m = 0
        does_line = False
    End If
End If

next_cmd_chk11:
If cmd_chk11(h).BackColor = 0 Then
    h_m = h_m + 1
    Label26(lbl_num).Caption = h_m
    does_line = True
    GoTo next_cmd_chk12
Else
    If does_line = True Then
        lbl_num = lbl_num + 1 'move next label
        h_m = 0
        does_line = False
    End If
End If

next_cmd_chk12:
If cmd_chk12(h).BackColor = 0 Then
    h_m = h_m + 1
    Label26(lbl_num).Caption = h_m
    does_line = True
    GoTo next_cmd_chk13
Else
    If does_line = True Then
        lbl_num = lbl_num + 1 'move next label
        h_m = 0
        does_line = False
    End If
End If

next_cmd_chk13:
If cmd_chk13(h).BackColor = 0 Then
    h_m = h_m + 1
    Label26(lbl_num).Caption = h_m
    does_line = True
    GoTo next_cmd_chk14
Else
    If does_line = True Then
        lbl_num = lbl_num + 1 'move next label
        h_m = 0
        does_line = False
    End If
End If

next_cmd_chk14:
If cmd_chk14(h).BackColor = 0 Then
    h_m = h_m + 1
    Label26(lbl_num).Caption = h_m
    does_line = True
    GoTo next_cmd_chk15
Else
    If does_line = True Then
        lbl_num = lbl_num + 1 'move next label
        h_m = 0
        does_line = False
    End If
End If

next_cmd_chk15:
If cmd_chk15(h).BackColor = 0 Then
    h_m = h_m + 1
    Label26(lbl_num).Caption = h_m
    does_line = True
    does_line = True
    GoTo next_cmd_chk16
Else
    If does_line = True Then
        lbl_num = lbl_num + 1 'move next label
        h_m = 0
        does_line = False
    End If
End If

next_cmd_chk16:
If cmd_chk16(h).BackColor = 0 Then
    h_m = h_m + 1
    Label26(lbl_num).Caption = h_m
    does_line = True
    GoTo next_cmd_chk17
Else
    If does_line = True Then
        lbl_num = lbl_num + 1 'move next label
        h_m = 0
        does_line = False
    End If
End If

next_cmd_chk17:
If cmd_chk17(h).BackColor = 0 Then
    h_m = h_m + 1
    Label26(lbl_num).Caption = h_m
    does_line = True
    GoTo next_cmd_chk18
Else
    If does_line = True Then
        lbl_num = lbl_num + 1 'move next label
        h_m = 0
        does_line = False
    End If
End If

next_cmd_chk18:
If cmd_chk18(h).BackColor = 0 Then
    h_m = h_m + 1
    Label26(lbl_num).Caption = h_m
    does_line = True
    GoTo next_cmd_chk19
Else
    If does_line = True Then
        lbl_num = lbl_num + 1 'move next label
        h_m = 0
        does_line = False
    End If
End If

next_cmd_chk19:
If cmd_chk19(h).BackColor = 0 Then
    h_m = h_m + 1
    Label26(lbl_num).Caption = h_m
    does_line = True
    GoTo next_cmd_chk20
Else
    If does_line = True Then
        lbl_num = lbl_num + 1 'move next label
        h_m = 0
        does_line = False
    End If
End If

next_cmd_chk20:
If cmd_chk20(h).BackColor = 0 Then
    h_m = h_m + 1
    Label26(lbl_num).Caption = h_m
    does_line = True
    GoTo next_cmd_chk21
Else
    If does_line = True Then
        lbl_num = lbl_num + 1 'move next label
        h_m = 0
        does_line = False
    End If
End If

next_cmd_chk21:
If cmd_chk21(h).BackColor = 0 Then
    h_m = h_m + 1
    Label26(lbl_num).Caption = h_m
    does_line = True
    GoTo next_cmd_chk22
Else
    If does_line = True Then
        lbl_num = lbl_num + 1 'move next label
        h_m = 0
        does_line = False
    End If
End If

next_cmd_chk22:
If cmd_chk22(h).BackColor = 0 Then
    h_m = h_m + 1
    Label26(lbl_num).Caption = h_m
    does_line = True
    GoTo next_cmd_chk23
Else
    If does_line = True Then
        lbl_num = lbl_num + 1 'move next label
        h_m = 0
        does_line = False
    End If
End If

next_cmd_chk23:
If cmd_chk23(h).BackColor = 0 Then
    h_m = h_m + 1
    Label26(lbl_num).Caption = h_m
    does_line = True
    GoTo next_cmd_chk24
Else
    If does_line = True Then
        lbl_num = lbl_num + 1 'move next label
        h_m = 0
        does_line = False
    End If
End If

next_cmd_chk24:
If cmd_chk24(h).BackColor = 0 Then
    h_m = h_m + 1
   ' If lbl_num = 300 Then ' !!!!!!!!!!!!!
    'lbl_num = lbl_num - 1
    'End If
    Label26(lbl_num).Caption = h_m
    does_line = True
    GoTo next_cmd_chk25
Else
    If does_line = True Then
        lbl_num = lbl_num + 1 'move next label
        h_m = 0
        does_line = False
    End If
End If

next_cmd_chk25:
If cmd_chk25(h).BackColor = 0 Then
    h_m = h_m + 1
    'If lbl_num = 300 Then ' !!!!!!!!!!!!!
    'lbl_num = lbl_num - 1
    'End If
    does_line = True
    Label26(lbl_num).Caption = h_m

Else
    If does_line = True Then
        lbl_num = lbl_num + 1 'move next label
        h_m = 0
        does_line = False
    End If
End If




Next

'sof calc LINES! #################################

'fixig labels LINES! ###################################################

Dim arr_fix_line(324, 1), line As Integer

For i = 0 To 324
arr_fix_line(i, 0) = 0
arr_fix_line(i, 1) = 0
Next

line = 0
again:
h_f = 0  ' Label1

If line = 325 Then
GoTo sof_fixig_labels_LINES
End If

For i = 0 To 12
    If Not (Label26(i + line).Caption = "") Then
        arr_fix_line(i + line, 0) = i + line ' index
        arr_fix_line(i + line, 1) = CStr(Label26(i + line).Caption)
        Label26(i + line).Caption = "" ' מאפס את הלייבל שלא רלוונטי יותר
        h_f = h_f + 1
    End If
Next

For j = 0 To h_f - 1 ' לתקוע למעלה
    Label26((12 - h_f + 1 + arr_fix_line(j + line, 0))).Caption = Str(arr_fix_line(j + line, 1))
Next

line = line + 13
GoTo again


'!For i = 0 To 324
'!arr_fix_line(i, 0) = 0
'!arr_fix_line(i, 1) = 0
'!Next

'For i = 0 To 324
 '   Label26(i).Caption = Str(arr_fix_line(i, 1))
'Next


'!For j = 0 To h_f - 1 ' לתקוע למעלה
'!    Label26((12 - h_f + 1 + arr_fix_line(j, 0))).Caption = Str(arr_fix_line(j, 1))
'! '    Label1(arr_fix(j, 0)).Caption = "" האיפוס למעלה
'!Next

sof_fixig_labels_LINES:
'sof fixig labels LINES! ###################################################

    
End Sub

Sub marking_line(num) 'marking_line #################################

cmd_chk1(num).BackColor = 0
cmd_chk2(num).BackColor = 0
cmd_chk3(num).BackColor = 0
cmd_chk4(num).BackColor = 0
cmd_chk5(num).BackColor = 0
cmd_chk6(num).BackColor = 0
cmd_chk7(num).BackColor = 0
cmd_chk8(num).BackColor = 0
cmd_chk9(num).BackColor = 0
cmd_chk10(num).BackColor = 0
cmd_chk11(num).BackColor = 0
cmd_chk12(num).BackColor = 0
cmd_chk13(num).BackColor = 0
cmd_chk14(num).BackColor = 0
cmd_chk15(num).BackColor = 0
cmd_chk16(num).BackColor = 0
cmd_chk17(num).BackColor = 0
cmd_chk18(num).BackColor = 0
cmd_chk19(num).BackColor = 0
cmd_chk20(num).BackColor = 0
cmd_chk21(num).BackColor = 0
cmd_chk22(num).BackColor = 0
cmd_chk23(num).BackColor = 0
cmd_chk24(num).BackColor = 0
cmd_chk25(num).BackColor = 0
End Sub

Private Sub cmd_mrk_line_Click(Index As Integer)
Call marking_line(Index)
End Sub

Sub clearing_line(num) 'clearing_line #################################
cmd_chk1(num).BackColor = &HFFFFFF
cmd_chk2(num).BackColor = &HFFFFFF
cmd_chk3(num).BackColor = &HFFFFFF
cmd_chk4(num).BackColor = &HFFFFFF
cmd_chk5(num).BackColor = &HFFFFFF
cmd_chk6(num).BackColor = &HFFFFFF
cmd_chk7(num).BackColor = &HFFFFFF
cmd_chk8(num).BackColor = &HFFFFFF
cmd_chk9(num).BackColor = &HFFFFFF
cmd_chk10(num).BackColor = &HFFFFFF
cmd_chk11(num).BackColor = &HFFFFFF
cmd_chk12(num).BackColor = &HFFFFFF
cmd_chk13(num).BackColor = &HFFFFFF
cmd_chk14(num).BackColor = &HFFFFFF
cmd_chk15(num).BackColor = &HFFFFFF
cmd_chk16(num).BackColor = &HFFFFFF
cmd_chk17(num).BackColor = &HFFFFFF
cmd_chk18(num).BackColor = &HFFFFFF
cmd_chk19(num).BackColor = &HFFFFFF
cmd_chk20(num).BackColor = &HFFFFFF
cmd_chk21(num).BackColor = &HFFFFFF
cmd_chk22(num).BackColor = &HFFFFFF
cmd_chk23(num).BackColor = &HFFFFFF
cmd_chk24(num).BackColor = &HFFFFFF
cmd_chk25(num).BackColor = &HFFFFFF
End Sub
Private Sub cmd_clr_line_Click(Index As Integer)
Call clearing_line(Index)
End Sub



'############################################# marks tur


'Private Sub cmd_mrk_tur1_Click()

'If marks_tur1 = True Then
 '   For i = 0 To 24
  '      cmd_chk1(i).BackColor  = 0
  '  Next
  '  'cmd_mrk_tur1.Caption "X"
  '  marks_tur1 = False
'Else
 '   For i = 0 To 24
  '      cmd_chk1(i).BackColor  = 0
   ' Next
    'cmd_mrk_tur1.Caption ("V")
   ' marks_tur1 = True
'End If

'End Sub


Private Sub Command1_Click()
    For i = 0 To 24
        cmd_chk1(i).BackColor = 0
    Next
End Sub
Private Sub Command2_Click()
    For i = 0 To 24
        cmd_chk2(i).BackColor = 0
    Next
End Sub


Private Sub Command3_Click()
    For i = 0 To 24
        cmd_chk3(i).BackColor = 0
    Next
End Sub
Private Sub Command4_Click()
    For i = 0 To 24
        cmd_chk4(i).BackColor = 0
    Next
End Sub
Private Sub Command5_Click()
    For i = 0 To 24
        cmd_chk5(i).BackColor = 0
    Next
End Sub
Private Sub Command6_Click()
    For i = 0 To 24
        cmd_chk6(i).BackColor = 0
    Next
End Sub
Private Sub Command7_Click()
    For i = 0 To 24
        cmd_chk7(i).BackColor = 0
    Next
End Sub
Private Sub Command8_Click()
    For i = 0 To 24
        cmd_chk8(i).BackColor = 0
    Next
End Sub
Private Sub Command9_Click()
    For i = 0 To 24
        cmd_chk9(i).BackColor = 0
    Next
End Sub
Private Sub Command10_Click()
    For i = 0 To 24
        cmd_chk10(i).BackColor = 0
    Next
End Sub
Private Sub Command11_Click()
    For i = 0 To 24
        cmd_chk11(i).BackColor = 0
    Next
End Sub
Private Sub Command12_Click()
    For i = 0 To 24
        cmd_chk12(i).BackColor = 0
    Next
End Sub
Private Sub Command13_Click()
    For i = 0 To 24
        cmd_chk13(i).BackColor = 0
    Next
End Sub
Private Sub Command14_Click()
    For i = 0 To 24
        cmd_chk14(i).BackColor = 0
    Next
End Sub
Private Sub Command15_Click()
    For i = 0 To 24
        cmd_chk15(i).BackColor = 0
    Next
End Sub
Private Sub Command16_Click()
    For i = 0 To 24
        cmd_chk16(i).BackColor = 0
    Next
End Sub
Private Sub Command17_Click()
    For i = 0 To 24
        cmd_chk17(i).BackColor = 0
    Next
End Sub
Private Sub Command18_Click()
    For i = 0 To 24
        cmd_chk18(i).BackColor = 0
    Next
End Sub
Private Sub Command19_Click()
    For i = 0 To 24
        cmd_chk19(i).BackColor = 0
    Next
End Sub
Private Sub Command20_Click()
    For i = 0 To 24
        cmd_chk20(i).BackColor = 0
    Next
End Sub
Private Sub Command21_Click()
    For i = 0 To 24
        cmd_chk21(i).BackColor = 0
    Next
End Sub
Private Sub Command22_Click()
    For i = 0 To 24
        cmd_chk22(i).BackColor = 0
    Next
End Sub
Private Sub Command23_Click()
    For i = 0 To 24
        cmd_chk23(i).BackColor = 0
    Next
End Sub
Private Sub Command24_Click()
    For i = 0 To 24
        cmd_chk24(i).BackColor = 0
    Next
End Sub
Private Sub Command25_Click()
    For i = 0 To 24
        cmd_chk25(i).BackColor = 0
    Next
End Sub

'############################################# clears tur


Private Sub Command26_Click()
    For i = 0 To 24
        cmd_chk25(i).BackColor = &HFFFFFF
    Next
End Sub
Private Sub Command27_Click()
    For i = 0 To 24
        cmd_chk24(i).BackColor = &HFFFFFF
    Next
End Sub


Private Sub Command28_Click()
    For i = 0 To 24
        cmd_chk23(i).BackColor = &HFFFFFF
    Next
End Sub
Private Sub Command29_Click()
    For i = 0 To 24
        cmd_chk22(i).BackColor = &HFFFFFF
    Next
End Sub
Private Sub Command30_Click()
    For i = 0 To 24
        cmd_chk21(i).BackColor = &HFFFFFF
    Next
End Sub
Private Sub Command31_Click()
    For i = 0 To 24
        cmd_chk20(i).BackColor = &HFFFFFF
    Next
End Sub
Private Sub Command32_Click()
    For i = 0 To 24
        cmd_chk19(i).BackColor = &HFFFFFF
    Next
End Sub
Private Sub Command33_Click()
    For i = 0 To 24
        cmd_chk18(i).BackColor = &HFFFFFF
    Next
End Sub
Private Sub Command34_Click()
    For i = 0 To 24
        cmd_chk17(i).BackColor = &HFFFFFF
    Next
End Sub
Private Sub Command35_Click()
    For i = 0 To 24
        cmd_chk16(i).BackColor = &HFFFFFF
    Next
End Sub
Private Sub Command36_Click()
    For i = 0 To 24
        cmd_chk15(i).BackColor = &HFFFFFF
    Next
End Sub
Private Sub Command37_Click()
    For i = 0 To 24
        cmd_chk14(i).BackColor = &HFFFFFF
    Next
End Sub
Private Sub Command38_Click()
    For i = 0 To 24
        cmd_chk13(i).BackColor = &HFFFFFF
    Next
End Sub
Private Sub Command39_Click()
    For i = 0 To 24
        cmd_chk12(i).BackColor = &HFFFFFF
    Next
End Sub
Private Sub Command40_Click()
    For i = 0 To 24
        cmd_chk11(i).BackColor = &HFFFFFF
    Next
End Sub
Private Sub Command41_Click()
    For i = 0 To 24
        cmd_chk10(i).BackColor = &HFFFFFF
    Next
End Sub
Private Sub Command42_Click()
    For i = 0 To 24
        cmd_chk9(i).BackColor = &HFFFFFF
    Next
End Sub
Private Sub Command43_Click()
    For i = 0 To 24
        cmd_chk8(i).BackColor = &HFFFFFF
    Next
End Sub
Private Sub Command44_Click()
    For i = 0 To 24
        cmd_chk7(i).BackColor = &HFFFFFF
    Next
End Sub
Private Sub Command45_Click()
    For i = 0 To 24
        cmd_chk6(i).BackColor = &HFFFFFF
    Next
End Sub
Private Sub Command46_Click()
    For i = 0 To 24
        cmd_chk5(i).BackColor = &HFFFFFF
    Next
End Sub
Private Sub Command47_Click()
    For i = 0 To 24
        cmd_chk4(i).BackColor = &HFFFFFF
    Next
End Sub
Private Sub Command48_Click()
    For i = 0 To 24
        cmd_chk3(i).BackColor = &HFFFFFF
    Next
End Sub
Private Sub Command49_Click()
    For i = 0 To 24
        cmd_chk2(i).BackColor = &HFFFFFF
    Next
End Sub
Private Sub Command50_Click()
    For i = 0 To 24
        cmd_chk1(i).BackColor = &HFFFFFF
    Next
End Sub


Private Sub Form_Load()
For i = 1 To 24

cmd_chk4(i).Top = cmd_chk2(i).Top
cmd_chk4(i).Left = cmd_chk4(0).Left

cmd_chk3(i).Top = cmd_chk2(i).Top
cmd_chk3(i).Left = cmd_chk3(0).Left

cmd_chk5(i).Top = cmd_chk2(i).Top
cmd_chk5(i).Left = cmd_chk5(0).Left

cmd_chk6(i).Top = cmd_chk2(i).Top
cmd_chk6(i).Left = cmd_chk6(0).Left

cmd_chk7(i).Top = cmd_chk2(i).Top
cmd_chk7(i).Left = cmd_chk7(0).Left

cmd_chk8(i).Top = cmd_chk2(i).Top
cmd_chk8(i).Left = cmd_chk8(0).Left

cmd_chk9(i).Top = cmd_chk2(i).Top
cmd_chk9(i).Left = cmd_chk9(0).Left

cmd_chk10(i).Top = cmd_chk2(i).Top
cmd_chk10(i).Left = cmd_chk10(0).Left

cmd_chk11(i).Top = cmd_chk2(i).Top
cmd_chk11(i).Left = cmd_chk11(0).Left

cmd_chk12(i).Top = cmd_chk2(i).Top
cmd_chk12(i).Left = cmd_chk12(0).Left

cmd_chk13(i).Top = cmd_chk2(i).Top
cmd_chk13(i).Left = cmd_chk13(0).Left

cmd_chk14(i).Top = cmd_chk2(i).Top
cmd_chk14(i).Left = cmd_chk14(0).Left

cmd_chk15(i).Top = cmd_chk2(i).Top
cmd_chk15(i).Left = cmd_chk15(0).Left

cmd_chk16(i).Top = cmd_chk2(i).Top
cmd_chk16(i).Left = cmd_chk16(0).Left

cmd_chk17(i).Top = cmd_chk2(i).Top
cmd_chk17(i).Left = cmd_chk17(0).Left

cmd_chk18(i).Top = cmd_chk2(i).Top
cmd_chk18(i).Left = cmd_chk18(0).Left

cmd_chk19(i).Top = cmd_chk2(i).Top
cmd_chk19(i).Left = cmd_chk19(0).Left

cmd_chk20(i).Top = cmd_chk2(i).Top
cmd_chk20(i).Left = cmd_chk20(0).Left

cmd_chk21(i).Top = cmd_chk2(i).Top
cmd_chk21(i).Left = cmd_chk21(0).Left

cmd_chk22(i).Top = cmd_chk2(i).Top
cmd_chk22(i).Left = cmd_chk22(0).Left

cmd_chk23(i).Top = cmd_chk2(i).Top
cmd_chk23(i).Left = cmd_chk23(0).Left

cmd_chk24(i).Top = cmd_chk2(i).Top
cmd_chk24(i).Left = cmd_chk24(0).Left

cmd_chk25(i).Top = cmd_chk2(i).Top
cmd_chk25(i).Left = cmd_chk25(0).Left




'set  Form1.cmd_chk3(i)
'Set cmd_chkboxes(i, j) = Form1.Controls.Add("VB.CommandButton", "CommandButton" + Trim$(Str(i * 100 + j)), Form1)
'Set cmd_chk4(i) = Form1.Controls.Add("VB.CommandButton", "CommandButton" + Trim$(Str(i * 100)), Form1)

Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim intResponse As Integer
intResponse = MsgBox("Don't you want to save before quitting ?", vbYesNoCancel + vbQuestion, "Quit")
If intResponse = vbYes Then
Call cmdSave_Click
ElseIf intResponse = vbNo Then
    End
Else
    Cancel = 1
End If

End Sub

Private Sub tmr_rnd_Timer()

Randomize

For i = 0 To 24


If Int((2 * Rnd)) = 1 Then
cmd_chk1(i).BackColor = QBColor(15)
Else
cmd_chk1(i).BackColor = 0
End If


If Int((2 * Rnd)) = 1 Then
cmd_chk2(i).BackColor = QBColor(15)
Else
cmd_chk2(i).BackColor = 0
End If


If Int((2 * Rnd)) = 1 Then
cmd_chk3(i).BackColor = QBColor(15)
Else
cmd_chk3(i).BackColor = 0
End If


If Int((2 * Rnd)) = 1 Then
cmd_chk4(i).BackColor = QBColor(15)
Else
cmd_chk4(i).BackColor = 0
End If


If Int((2 * Rnd)) = 1 Then
cmd_chk5(i).BackColor = QBColor(15)
Else
cmd_chk5(i).BackColor = 0
End If


If Int((2 * Rnd)) = 1 Then
cmd_chk6(i).BackColor = QBColor(15)
Else
cmd_chk6(i).BackColor = 0
End If


If Int((2 * Rnd)) = 1 Then
cmd_chk7(i).BackColor = QBColor(15)
Else
cmd_chk7(i).BackColor = 0
End If


If Int((2 * Rnd)) = 1 Then
cmd_chk8(i).BackColor = QBColor(15)
Else
cmd_chk8(i).BackColor = 0
End If


If Int((2 * Rnd)) = 1 Then
cmd_chk9(i).BackColor = QBColor(15)
Else
cmd_chk9(i).BackColor = 0
End If


If Int((2 * Rnd)) = 1 Then
cmd_chk10(i).BackColor = QBColor(15)
Else
cmd_chk10(i).BackColor = 0
End If


If Int((2 * Rnd)) = 1 Then
cmd_chk11(i).BackColor = QBColor(15)
Else
cmd_chk11(i).BackColor = 0
End If


If Int((2 * Rnd)) = 1 Then
cmd_chk12(i).BackColor = QBColor(15)
Else
cmd_chk12(i).BackColor = 0
End If

If Int((2 * Rnd)) = 1 Then
cmd_chk13(i).BackColor = QBColor(15)
Else
cmd_chk13(i).BackColor = 0
End If


If Int((2 * Rnd)) = 1 Then
cmd_chk14(i).BackColor = QBColor(15)
Else
cmd_chk14(i).BackColor = 0
End If


If Int((2 * Rnd)) = 1 Then
cmd_chk15(i).BackColor = QBColor(15)
Else
cmd_chk15(i).BackColor = 0
End If


If Int((2 * Rnd)) = 1 Then
cmd_chk16(i).BackColor = QBColor(15)
Else
cmd_chk16(i).BackColor = 0
End If


If Int((2 * Rnd)) = 1 Then
cmd_chk17(i).BackColor = QBColor(15)
Else
cmd_chk17(i).BackColor = 0
End If


If Int((2 * Rnd)) = 1 Then
cmd_chk18(i).BackColor = QBColor(15)
Else
cmd_chk18(i).BackColor = 0
End If


If Int((2 * Rnd)) = 1 Then
cmd_chk19(i).BackColor = QBColor(15)
Else
cmd_chk19(i).BackColor = 0
End If


If Int((2 * Rnd)) = 1 Then
cmd_chk20(i).BackColor = QBColor(15)
Else
cmd_chk20(i).BackColor = 0
End If


If Int((2 * Rnd)) = 1 Then
cmd_chk21(i).BackColor = QBColor(15)
Else
cmd_chk21(i).BackColor = 0
End If


If Int((2 * Rnd)) = 1 Then
cmd_chk22(i).BackColor = QBColor(15)
Else
cmd_chk22(i).BackColor = 0
End If


If Int((2 * Rnd)) = 1 Then
cmd_chk23(i).BackColor = QBColor(15)
Else
cmd_chk23(i).BackColor = 0
End If


If Int((2 * Rnd)) = 1 Then
cmd_chk24(i).BackColor = QBColor(15)
Else
cmd_chk24(i).BackColor = 0
End If

If Int((2 * Rnd)) = 1 Then
cmd_chk25(i).BackColor = QBColor(15)
Else
cmd_chk25(i).BackColor = 0
End If

Next
End Sub
