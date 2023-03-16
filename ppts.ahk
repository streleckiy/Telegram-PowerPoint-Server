#SingleInstance force
#NoEnv
#NoTrayIcon
#If isMainWindowActive()

onError("error")
OnExit("exit")

refreshControlMessage() {
   global

   try {
      slideNumber := ppSlideRunView.Slide.SlideNumber
      slidesCount := ppObj.ActivePresentation.Slides.Count
   }
   catch e {
      try ppSlideRunView.Exit
      sendTelegramRequest("editMessageText", "chat_id=" owner_id "&message_id=" last_msg_id "&text=<b>Ожидание.</b>`n`nОжидается окно Microsoft PowerPoint в режиме редактирования слайдов для подключения к его API...&parse_mode=html")
      initPPObject()
      return
   }
   
   sendTelegramRequest("editMessageText", "chat_id=" owner_id "&message_id=" last_msg_id "&parse_mode=html&text=<b>Управление презентацией.</b>`n`n<b>Сейчас производится слайд: </b><code>№" slideNumber "</code> из <code>" slidesCount "</code>.&reply_markup=" keyboard)
}

initPPObject() {
   global ppObj, ppSlideShowSettings, owner_id, last_msg_id
   ppObj := ""

   while (!IsObject(ppObj)) {
      try ppObj := ComObjActive("PowerPoint.Application")
      catch e {
         sleep 100
         continue
      }
   }
         
   loop {
      try ppSlideShowSettings := ppObj.ActivePresentation.SlideShowSettings
      catch e {
         sleep 100
         continue
      }
            
      break
   }

	keyboard = {"inline_keyboard":[[{"text":"\ud83e\udde9 \u041f\u043e\u0434\u043a\u043b\u044e\u0447\u0438\u0442\u044c\u0441\u044f \u043a \u043f\u0440\u0435\u0437\u0435\u043d\u0442\u0430\u0446\u0438\u0438","callback_data":"startPresentation"}]]}
	sendTelegramRequest("editMessageText", "chat_id=" owner_id "&message_id=" last_msg_id "&parse_mode=html&text=<b>Готов к началу презентации.</b>`n`nНажимайте на кнопки для управления.&reply_markup=" keyboard)
}

exit() {
   global last_msg_id, owner_id
   if (last_msg_id)
      sendTelegramRequest("editMessageText", "chat_id=" owner_id "&message_id=" last_msg_id "&parse_mode=html&text=<b>Сессия завершена.</b>`n`nБот отключен.")
}

error(e) {
   MsgBox, 16, % title, Произошла критическая ошибка. Программа перезагрузится через 3 секунды., 3
   sleep 3000
   reload
   sleep 10000
}

class JSON
{
   static JS := JSON._GetJScriptObject(), true := {}, false := {}, null := {}
   
   Parse(sJson, js := false)  {
      if jsObj := this.VerifyJson(sJson)
         Return js ? jsObj : this._CreateObject(jsObj)
   }
   
   Stringify(obj, js := false, indent := "") {
      if (js && !RegExMatch(js, "\s"))
         Return this.JS.JSON.stringify(obj, "", indent)
      else {
         (RegExMatch(js, "\s") && indent := js)
         sObj := this._ObjToString(obj)
         Return this.JS.eval("JSON.stringify(" . sObj . ",'','" . indent . "')")
      }
   }
   
   GetKey(sJson, key, indent := "") {
	  if !this.VerifyJson(sJson)
         Return
		 
		 symbol = `"
      try Return StrReplace(StrReplace(StrReplace(Ltrim(RTrim(this.JS.eval("JSON.stringify((" . sJson . ")" . (SubStr(key, 1, 1) = "[" ? "" : ".") . key . ",'','" . indent . "')"), symbol), symbol), "\/", "/"), "\n", "`n"), "\" symbol, symbol)
      catch
         console.writeln("[DEBUG] " StrReplace(RSHELL_JSON_TEXT_BAD_KEY_TO_DEBUG, "%1", key))
   }
   
   SetKey(sJson, key, value, indent := "") {
      if !this.VerifyJson(sJson)
         Return
      if !this.VerifyJson(value, true) {
         console.warning(StrReplace(RSHELL_JSON_TEXT_BAD_VALUE, "%1", value))
         Return
      }
      try {
         res := this.JS.eval( "var obj = (" . sJson . ");"
                            . "obj" . (SubStr(key, 1, 1) = "[" ? "" : ".") . key . "=" . value . ";"
                            . "JSON.stringify(obj,'','" . indent . "')" )
         this.JS.eval("obj = ''")
         Return res
      }
      catch
         console.writeln("[DEBUG] " StrReplace(RSHELL_JSON_TEXT_BAD_KEY_TO_DEBUG, "%1", key))
   }
   
   RemoveKey(sJson, key, indent := "") {
      if !this.VerifyJson(sJson)
         Return
      
      sign := SubStr(key, 1, 1) = "[" ? "" : "."
      try {
         if !RegExMatch(key, "(.*)\[(\d+)]$", match)
            res := this.JS.eval("var obj = (" . sJson . "); delete obj" . sign . key . "; JSON.stringify(obj,'','" . indent . "')")
         else
            res := this.JS.eval( "var obj = (" . sJson . ");" 
                               . "obj" . (match1 != "" ? sign . match1 : "") . ".splice(" . match2 . ", 1);"
                               . "JSON.stringify(obj,'','" . indent . "')" )
         this.JS.eval("obj = ''")
         Return res
      }
      catch
         console.writeln("[DEBUG] " StrReplace(RSHELL_JSON_TEXT_BAD_KEY_TO_DEBUG, "%1", key))
   }
   
   Enum(sJson, key := "", indent := "") {
      if !this.VerifyJson(sJson)
         Return
      
      conc := key ? (SubStr(key, 1, 1) = "[" ? "" : ".") . key : ""
      try {
         jsObj := this.JS.eval("(" sJson ")" . conc)
         res := jsObj.IsArray()
         if (res = "")
            Return
         obj := {}
         if (res = -1) {
            Loop % jsObj.length
               obj[A_Index - 1] := this.JS.eval("JSON.stringify((" sJson ")" . conc . "[" . (A_Index - 1) . "],'','" . indent . "')")
         }
         else if (res = 0) {
            keys := jsObj.GetKeys()
            Loop % keys.length
               k := keys[A_Index - 1], obj[k] := this.JS.eval("JSON.stringify((" sJson ")" . conc . "['" . k . "'],'','" . indent . "')")
         }
         Return obj
      }
      catch
         console.writeln("[DEBUG] " StrReplace(RSHELL_JSON_TEXT_BAD_KEY_TO_DEBUG, "%1", key))
   }
   
   VerifyJson(sJson, silent := false) {
      try jsObj := this.JS.eval("(" sJson ")")
      catch {
         if !silent
            console.writeln("[DEBUG] " StrReplace(RSHELL_JSON_TEXT_BAD_JSON_STRING_TO_DEBUG, "%1", sJson))
         Return
      }
      Return IsObject(jsObj) ? jsObj : true
   }
   
   _ObjToString(obj) {
      if IsObject( obj ) {
         for k, v in ["true", "false", "null"]
            if (obj = this[v])
               Return v
            
         isArray := true
         for key in obj {
            if IsObject(key)
               throw Exception("Invalid key")
            if !( key = A_Index || isArray := false )
               break
         }
         for k, v in obj
            str .= ( A_Index = 1 ? "" : "," ) . ( isArray ? "" : """" . k . """:" ) . this._ObjToString(v)

         Return isArray ? "[" str "]" : "{" str "}"
      }
      else if !(obj*1 = "" || RegExMatch(obj, "\s"))
         Return obj
      
      for k, v in [["\", "\\"], [A_Tab, "\t"], ["""", "\"""], ["/", "\/"], ["`n", "\n"], ["`r", "\r"], [Chr(12), "\f"], [Chr(08), "\b"]]
         obj := StrReplace( obj, v[1], v[2] )

      Return """" obj """"
   }

   _GetJScriptObject() {
      static doc
      doc := ComObjCreate("htmlfile")
      doc.write("<meta http-equiv=""X-UA-Compatible"" content=""IE=9"">")
      JS := doc.parentWindow
      JSON._AddMethods(JS)
      Return JS
   }

   _AddMethods(ByRef JS) {
      JScript =
      (
         Object.prototype.GetKeys = function () {
            var keys = []
            for (var k in this)
               if (this.hasOwnProperty(k))
                  keys.push(k)
            return keys
         }
         Object.prototype.IsArray = function () {
            var toStandardString = {}.toString
            return toStandardString.call(this) == '[object Array]'
         }
      )
      JS.eval(JScript)
   }

   _CreateObject(jsObj) {
      res := jsObj.IsArray()
      if (res = "")
         Return jsObj
      
      else if (res = -1) {
         obj := []
         Loop % jsObj.length
            obj[A_Index] := this._CreateObject(jsObj[A_Index - 1])
      }
      else if (res = 0) {
         obj := {}
         keys := jsObj.GetKeys()
         Loop % keys.length
            k := keys[A_Index - 1], obj[k] := this._CreateObject(jsObj[k])
      }
      Return obj
   }
}

isMainWindowActive() {
	IfWinActive, ahk_id %mainwid%
		return true
	else
		return false
}

sendTelegramRequest(method, content := "") {
	global telegramToken

	try whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	try whr.Open("POST", "https://api.telegram.org/bot" telegramToken "/" method, true)
	try whr.SetRequestHeader("User-Agent", "Renux Shell v" version)
	try whr.SetRequestHeader("Content-Type","application/x-www-form-urlencoded")
	try whr.Send(content)
	try whr.WaitForResponse()
	try response := whr.ResponseText
   catch e {
      MsgBox, 16, % title, Ошибка при отправке запроса на сервер Telegram. Возможно`, у Вас нестабильное подключение к интернету., 3
   }

	return response
}

global title := "vk.com/strdev - PowerPoint Telegram Server"
global action := 0
global mainwid
global telegramToken

IniRead, telegramToken, config.ini, tg, token, % " "

Gui, +hwndmainwid +OwnDialogs
Gui, Color, White
Gui, Font, S14 CDefault, Segoe UI
Gui, Add, Text, x12 y7 w451 h25 , Добро пожаловать
Gui, Font, S10 CDefault, Segoe UI
Gui, Add, Progress, x-7 y40 w489 h1 -Border +cGray, 100
Gui, Add, Text, x12 y45 w451 h19 , Получите токен бота от @BotFather и вставьте его сюда.
Gui, Font, S9 CDefault, Segoe UI
Gui, Add, Edit, x12 y75 w451 h19 vInputedTelegramToken, % telegramToken
Gui, Add, Button, x12 y107 w134 h28 gOpenDeveloperPage, Разработчик
Gui, Add, Button, x300 y107 w163 h28 gInputToken, Продолжить (Enter)
Gui, Show, w479 h145, % title
return

~Enter::
if (action == 0)
	goto InputToken

return

OpenDeveloperPage:
MsgBox, 1, % title, Ссылка vk.com/strdev откроется в браузере по-умолчанию.
IfMsgBox, Ok
	Run, http://vk.com/strdev,, UseErrorLevel

return

GuiEscape:
GuiClose:
ExitApp
return

InputToken:
Gui, Submit, NoHide
telegramToken := InputedTelegramToken
response := sendTelegramRequest("getMe")

if (trim(telegramToken) == "")
   return

if (response == "")
   return

if (JSON.GetKey(response, "ok") == "false") {
	MsgBox, 16, % title, Вы указали некорректный токен.
	return
}

IniWrite % telegramToken, config.ini, tg, token

Gui, Destroy
Random, code, 1000, 9999
isWaitingCode = 1

Menu, Tray, Add, Отключить сервер, GuiClose
Menu, Tray, NoStandard
Menu, Tray, Icon
Menu, Tray, Tip, % title

TrayTip, % title, % "Вы ввели токен бота '" JSON.GetKey(response, "result.first_name") "' (@" JSON.GetKey(response, "result.username") ").`n`nОтправьте боту код " code " чтобы привязать аккаунт."

loop {
	response := sendTelegramRequest("getUpdates", "timeout=15&offset=" update_id+1)
	
	if (JSON.GetKey(response, "ok") == "false")
		continue

	update_id := JSON.GetKey(response, "result[0].update_id")

	if (isWaitingCode) {
		if (JSON.GetKey(response, "result[0].message.from.is_bot") == true) {
			continue
		}
		
		if (JSON.GetKey(response, "result[0].message.text") == code) {
			isWaitingCode = 0
			owner_id := JSON.GetKey(response, "result[0].message.chat.id")
			
			sendTelegramRequest("sendMessage", "chat_id=" owner_id "&text=<b>" title " успешно привязан к Вашему аккаунту.</b>`n`nСообщения от других пользователей <b>не будут</b> обрабатываться ботом.&parse_mode=html")
			response := sendTelegramRequest("sendMessage", "chat_id=" owner_id "&text=<b>Ожидание.</b>`n`nОжидается окно Microsoft PowerPoint в режиме редактирования слайдов для подключения к его API...&parse_mode=html")
			last_msg_id := JSON.GetKey(response, "result.message_id")
         TrayTip, % title, Успешно подключено. Используйте меню в трее для управления приложением.
         initPPObject()
         continue
		}
	}

	callback_data := "", callback_data := JSON.GetKey(response, "result[0].callback_query.data")
	callback_id := "", callback_id := JSON.GetKey(response, "result[0].callback_query.id")

	if (callback_data) {
		if (JSON.GetKey(response, "result[0].callback_query.from.id") != owner_id)
			continue	
	} else {
		if (JSON.GetKey(response, "result[0].message.chat.id") != owner_id) {
			continue
		}
	}

	; =======================================================

   if (callback_data == "startPresentation") {
      sendTelegramRequest("editMessageText", "chat_id=" owner_id "&message_id=" last_msg_id "&parse_mode=html&text=<b>Запуск презентации...</b>`n`nПодождите, пожалуйста...")
      try ppSlideShowSettings := ppObj.ActivePresentation.SlideShowSettings
      catch e {
         sendtelegramRequest("answerCallbackQuery", "callback_query_id=" callback_id "&text=Не удалось получить объект презентации!&show_alert=1&cache_time=0")
         sendTelegramRequest("editMessageText", "chat_id=" owner_id "&message_id=" last_msg_id "&text=<b>Ожидание.</b>`n`nОжидается окно Microsoft PowerPoint в режиме редактирования слайдов для подключения к его API...&parse_mode=html")
         initPPObject()
         continue
      }

      try ppSlideRunView := ppSlideShowSettings.Run.View
      catch e {
         sendtelegramRequest("answerCallbackQuery", "callback_query_id=" callback_id "&text=Не удалось подключиться к презентации!&show_alert=1&cache_time=0")
         sendTelegramRequest("editMessageText", "chat_id=" owner_id "&message_id=" last_msg_id "&text=<b>Ожидание.</b>`n`nОжидается окно Microsoft PowerPoint в режиме редактирования слайдов для подключения к его API...&parse_mode=html")
         initPPObject()
         continue
      }

      keyboard = {"inline_keyboard":[[{"text":"\u23ed \u0421\u043b\u0435\u0434\u0443\u044e\u0449\u0438\u0439 \u0441\u043b\u0430\u0439\u0434","callback_data":"nSlide"}],[{"text":"\u23ee \u041f\u0440\u0435\u0434\u044b\u0434\u0443\u0449\u0438\u0439 \u0441\u043b\u0430\u0439\u0434","callback_data":"pSlide"}],[{"text":"\ud83d\udd01 \u041f\u0435\u0440\u0435\u043f\u043e\u0434\u043a\u043b\u044e\u0447\u0438\u0442\u044c\u0441\u044f","callback_data":"startPresentation"}],[{"text":"\ud83d\udcf4 \u0417\u0430\u0432\u0435\u0440\u0448\u0438\u0442\u044c \u0441\u043b\u0430\u0439\u0434-\u0448\u043e\u0443","callback_data":"endPresentation"}]]}
      sendTelegramRequest("answerCallbackQuery", "callback_query_id=" callback_id "&text=✅ Слайд-шоу запущено. Подключение установлено.&show_alert=0&cache_time=0")
      refreshControlMessage()
      continue
   }

   if (callback_data == "nSlide") {
      try ppSlideRunView.Next
      catch e {
         sendTelegramRequest("answerCallbackQuery", "callback_query_id=" callback_id "&text=Не удалось переключить слайд на следующий!&show_alert=1&cache_time=0")
         continue
      }

      sendTelegramRequest("answerCallbackQuery", "callback_query_id=" callback_id "&text=✅ Готово.&show_alert=0&cache_time=0")
      refreshControlMessage()
      continue
   }

   if (callback_data == "pSlide") {
      try ppSlideRunView.Previous
      catch e {
         sendTelegramRequest("answerCallbackQuery", "callback_query_id=" callback_id "&text=Не удалось переключить слайд на предыдущий!&show_alert=1&cache_time=0")
         continue
      }

      sendTelegramRequest("answerCallbackQuery", "callback_query_id=" callback_id "&text=✅ Готово.&show_alert=0&cache_time=0")
      refreshControlMessage()
      continue
   }

   if (callback_data == "endPresentation") {
      sendTelegramRequest("answerCallbackQuery", "callback_query_id=" callback_id "&text=✅ Слайд-шоу остановлено.&show_alert=0&cache_time=0")
      try ppSlideRunView.Exit
      refreshControlMessage()
      continue
   }
}
return