Attribute VB_Name = "ASUS_MACHINE_EDIT_SWAP"
'THIS CHUNK MOVED TO TOP





Sub NPN()


      
        xsop = 0
        '------------------------------------
        If Menu.CheckMumBT.Value = vbChecked Then
            xxz = 0
            If InStr(Gcw$, "Login - BT Yahoo!") > 0 Then xxz = 1
            If xxz = 1 Then
                xsop = 1
                SendKeys "linda.lancaster@btinternet.com{tab}", 0
                SendKeys "grabmethesnatch2{enter}", 0
                ExeGo = True
                
            End If
        End If
      
        '------------------------------------
        If Menu.CheckEddieBT.Value = vbChecked Then
            xxz = 0
            If InStr(Gcw$, "Login - BT Yahoo!") > 0 Then xxz = 1
            If xxz = 1 Then
                xsop = 1
                SendKeys "edward.lancaster@btinternet.com{tab}", 0
                SendKeys "sunshine1{enter}", 0
                ExeGo = True
            
            End If
        End If
      
        'MINE
        '------------------------------------
        If xsop = 0 Then
            xxz = 0
            If InStr(Gcw$, "Login - BT Yahoo!") > 0 Then xxz = 1
            If xxz = 1 Then
                xsop = 1
                SendKeys "mosheraboveall2"
                ExeGo = True
                
            End If
        End If
            
        'https://bt.edit.client.yahoo.com/bt/create_sub?.partner=bt-1&.master_login=matthew.lancaster%40btopenworld.com&.done=%2Faccounts&.scrumb=sXi6wDhX1Gq&.intl=uk
        If xsop = 0 Then
            xxz = 0
            If InStr(URL$, "https://bt.edit.client.yahoo.com") > 0 Then xxz = 1
            If xxz = 1 Then
                xsop = 1
                SendKeys "mosheraboveall2"
                ExeGo = True
                
            End If
        End If
            
        xxz = 0
        If InStr(Gcw$, "Verify Password") > 0 Then xxz = 1
        If xxz = 1 Then
            SendKeys "mosheraboveall2"
            ExeGo = True
        
            dse = 1
        End If
        
        


End Sub

              
              
              
        
Sub ASUS_MACHINE_EDIT_SWAP()

        
        'https://bt.edit.client.yahoo.com/bt/create_sub?.partner=bt-1&.master_login=matthew.lancaster%40btopenworld.com&.done=%2Faccounts&.scrumb=sXi6wDhX1Gq&.intl=uk
        If xsop = 0 Then
            xxz = 0
            If InStr(URL$, "https://bt.edit.client.yahoo.com") > 0 Then xxz = 1
            If xxz = 1 Then
                xsop = 1
                SendKeys "mosheraboveall2"
                ExeGo = True
                
            End If
        End If
            
End Sub

