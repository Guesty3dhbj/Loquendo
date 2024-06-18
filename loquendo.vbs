          Dim Message, Speak
          Message=InputBox("Introduzca su texto por favor","Loquendo")
          Set Speak=CreateObject("sapi.spvoice")
          Speak.Speak Message