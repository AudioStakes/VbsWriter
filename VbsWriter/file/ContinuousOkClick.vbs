val = MsgBox( "OK選んでね" ,vbOKOnly +vbSystemModal + vbExclamation , "ここにタイトル")
If val = vbOK  Then val = MsgBox( "もう一回OK選んでね" ,vbOKOnly +vbSystemModal + vbExclamation , "ここにタイトル")
If val = vbOK  Then val = MsgBox( "さらにもう一回OK選んでね" ,vbOKOnly +vbSystemModal + vbExclamation , "ここにタイトル")
If val = vbOK  Then val = MsgBox( "おしまい" ,vbOKOnly +vbSystemModal + vbExclamation , "ここにタイトル")