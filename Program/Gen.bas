Attribute VB_Name = "Module1"
'??????????????????????????????????????????????????????????
'**********************************************************
'Program Title: The Ludo Game  version 1.0.0
'Author:        Samuel Johnson A.c
'Date:          April to May 2008
'copyright:     © 2008
'Email:         stormsamany@yahoo.co.uk
'Phone:         +234 808 574 6108

'Last Updated:  2nd of June 2008 by the Author
'if you have any problem going through this code
'you can contact the Author through the email address



' All Right Reserved

'It considered a criminal offence if this Program
'in whole or in part is published under a diffrent
'name, title or any alteration which may affect the
'functionality of the game for better or worse
'of which such changes were made, with the aim of
'publishing or reproducing such altered copies
'without the prior permission of the Author
'in a stored audio format (e.g a casette or audio disc)

'??????????????????????????????????????????????????????????
'**********************************************************


Public Pnum As Integer, bgNum(71) As Integer, Pnam1 As String, Pnam2 As String, Pnam3 As String, Pnam4 As String
Public Ptyp1 As Integer, Ptyp2 As Integer, Ptyp3 As Integer, Ptyp4 As Integer, PnumX As Integer, Loozer As String
Public P1out As Integer, P2out As Integer, P3out As Integer, P4out As Integer, LooserNam As String
Public Ply1Hom As Integer, Ply2Hom As Integer, Ply3Hom As Integer, Ply4Hom As Integer
Public PlayerType As Integer, Paused As Boolean, Rotate As Integer, Dirxtry  As String
Public ArrayOfWinners(3) As String
Public Const P1Left = 5760
Public Const P2Left = 480
Public Const P3Left = 480
Public Const P4left = 5760
Public Const P12Top = 8400
Public Const P34Top = 480


Public Sub Verify(Turn As String)
'there is a bug that this verify sub takes care of
'sometimes (very rare) the program malfunctions, occurs mostly for
'computer  players i think it's a time factor (speed)
'the no of seeds on the field for a particular player increase to 5(could be more, but 5 is what i've xprnced)
'and sometimes decreases to 3 in this case this particular player will never win

'so what we are going to do is this, always ensure that there is only and only a total of 4 seeds for each player
'increment if lacking or decrement if more than

Dim Index As Integer, Count As Integer, Inside As Boolean, Field As Integer
Dim Seed As Integer, Outside As Boolean, SeedO As Integer
Count = 0
Inside = False
Outside = False
With Board
   For Index = 71 To 0 Step -1
       If .BG(Index).Tag = Turn Then
       Count = Count + bgNum(Index)
       Field = Index
       Outside = True
       End If
   Next
   
   Select Case Turn
   
          Case "P1"
            For Index = 0 To 3
               If .Ply1(Index).Tag = 0 Then
                  Count = Count + 1
                  Seed = Index
                  Inside = True
               Else
                  SeedO = Index
               End If
            Next
            Count = Count + Ply1Hom
            
            If Count > 4 Then
               'first let's see if there is a inside then we will remove that seed
               If Inside = True Then
                  .Ply1(Seed).Tag = 1 'take it out
                  .Ply1(Seed).Visible = False
               ElseIf Outside = True Then 'no seed is inside check out a seed on the field
                   .BG(Field).Tag = ""
                   .BG(Field).Picture = LoadPicture("")
                   bgNum(Field) = 0
               Else 'then this means ply4hom is greater than 4 ?
               
                    'well no seed inside ,no seed outside but ply4hom is > 4
                    'this may not cause any bug ,i don't see any
               
               End If
           ElseIf Count < 4 Then
                ' well increment the current no by 1
                 ' now there could be a short of 2  but we'll be incrementn 1 at a time
                 'the best place to add the seed will be in prison
                 'since there will always be a space there if count < 4
                 .Ply1(SeedO).Tag = 0
                 .Ply1(SeedO).Visible = True
          End If
            
          Case "P2"
            For Index = 0 To 3
               If .Ply2(Index).Tag = 0 Then
                  Count = Count + 1
                  Seed = Index
                  Inside = True
               Else
                  SeedO = Index
               End If
            Next
            Count = Count + Ply2Hom
            If Count > 4 Then
               'first let's see if there is a inside then we will remove that seed
               If Inside = True Then
                  .Ply2(Seed).Tag = 1 'take it out
                  .Ply2(Seed).Visible = False
               ElseIf Outside = True Then 'no seed is inside check out a seed on the field
                   .BG(Field).Tag = ""
                   .BG(Field).Picture = LoadPicture("")
                   bgNum(Field) = 0
               Else 'then this means ply4hom is greater than 4 ?
               
                    'well no seed inside ,no seed outside but ply4hom is > 4
                    'this may not cause any bug ,i don't see any
               
               End If
           ElseIf Count < 4 Then
                ' well increment the current no by 1
                 ' now there could be a short of 2  but we'll be incrementn 1 at a time
                 'the best place to add the seed will be in prison
                 'since there will always be a space there if count < 4
                 .Ply2(SeedO).Tag = 0
                 .Ply2(SeedO).Visible = True
          End If

            
            
          Case "P3"
            For Index = 0 To 3
               If .Ply3(Index).Tag = 0 Then
                  Count = Count + 1
                  Seed = Index
                  Inside = True
               Else
                  SeedO = Index
               End If
            Next
            Count = Count + Ply3Hom
            If Count > 4 Then
               'first let's see if there is a inside then we will remove that seed
               If Inside = True Then
                  .Ply3(Seed).Tag = 1 'take it out
                  .Ply3(Seed).Visible = False
               ElseIf Outside = True Then 'no seed is inside check out a seed on the field
                   .BG(Field).Tag = ""
                   .BG(Field).Picture = LoadPicture("")
                   bgNum(Field) = 0
               Else 'then this means ply4hom is greater than 4 ?
               
                    'well no seed inside ,no seed outside but ply4hom is > 4
                    'this may not cause any bug ,i don't see any
               
               End If
           ElseIf Count < 4 Then
                ' well increment the current no by 1
                 ' now there could be a short of 2  but we'll be incrementn 1 at a time
                 'the best place to add the seed will be in prison
                 'since there will always be a space there if count < 4
                 .Ply3(SeedO).Tag = 0
                 .Ply3(SeedO).Visible = True
          End If

          Case "P4"
            For Index = 0 To 3
               If .Ply4(Index).Tag = 0 Then
                  Count = Count + 1
                  Seed = Index
                  Inside = True
               Else
                  SeedO = Index
               End If
            Next
            Count = Count + Ply4Hom
            
            If Count > 4 Then
               'first let's see if there is a inside then we will remove that seed
               If Inside = True Then
                  .Ply4(Seed).Tag = 1 'take it out
                  .Ply4(Seed).Visible = False
               ElseIf Outside = True Then 'no seed is inside check out a seed on the field
                   .BG(Field).Tag = ""
                   .BG(Field).Picture = LoadPicture("")
                   bgNum(Field) = 0
               Else 'then this means ply4hom is greater than 4 ?
               
                    'well no seed inside ,no seed outside but ply4hom is > 4
                    'this may not cause any bug ,i don't see any
               
               End If
           ElseIf Count < 4 Then
                ' well increment the current no by 1
                 ' now there could be a short of 2  but we'll be incrementn 1 at a time
                 'the best place to add the seed will be in prison
                 'since there will always be a space there if count < 4
                 .Ply4(SeedO).Tag = 0
                 .Ply4(SeedO).Visible = True
          End If
    End Select

End With


End Sub






