; This is free and unencumbered software released into the public domain.
;
; Anyone is free to copy, modify, publish, use, compile, sell, or
; distribute this software, either in source code form or as a compiled
; binary, for any purpose, commercial or non-commercial, and by any
; means.
;
; In jurisdictions that recognize copyright laws, the author or authors
; of this software dedicate any and all copyright interest in the
; software to the public domain. We make this dedication for the benefit
; of the public at large and to the detriment of our heirs and
; successors. We intend this dedication to be an overt act of
; relinquishment in perpetuity of all present and future rights to this
; software under copyright law.
;
; THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
; EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
; MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
; IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR
; OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
; ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
; OTHER DEALINGS IN THE SOFTWARE.
;
; For more information, please refer to <http://unlicense.org/>

#include <Constants.au3>
#include <Misc.au3>
#include <Date.au3>

; Налоговый период.
Global Const $g_iYear = 2016
; Исполняемый файл программы Декларация.
Global Const $g_sDeclBinary = "Decl" & String($g_iYear) & ".exe"
; Полный путь к программе.
Global Const $g_sPath = StringFormat("%s\Декларация %d\%s", @ProgramFilesDir, $g_iYear, $g_sDeclBinary)
; Заголовок основного окна программы.
Global Const $g_sTitleMain = "Без имени - Декларация " & String($g_iYear)
; Координаты кнопки "Доходы за пределами РФ".
Global Const $g_iXxRus = 98, $g_iYxRus = 414
; Координаты кнопки "Добавить источник выплат".
Global Const $g_ixAddSrc = 206, $g_iYaddSrc = 113
; Источник выплаты дохода по умолчанию.
Global Const $g_sDefaultIncomeSource = "Брокер ""XXX"" (США)"
; Код страны по умолчанию (США).
Global Const $g_iDefaultCountryCode = 840
; Код валюты по умолчанию (доллар США).
Global Const $g_iDefaultCurrencyCode = 840
; Код дохода "Дивиденды".
Global Const $g_iIncomeTypeDividends = 1010

; Подготовка к вводу данных:
; запускаю программу,
; отмечаю "Имеются доходы - В иностранной валюте",
; перехожу в раздел "Доходы за пределами РФ",
; закрываю окно с сообщением "Не введен номер инспекции".
OpenDecl()

; Ввод данных о дивидендах:
; дата, поступившая сумма, удержанная брокером сумма, описание бумаги.
; Дата в формате YYYY-MM-DD т.е. ГГГГ-ММ-ДД т.е. ГОД-МЕСЯЦ-ДЕНЬ .
$sDesc = "ISHARES CORE US AGGREGATE BOND ETF"
FillDiv("2016-02-01", 22.8, 2.28, $sDesc)
; ...
FillDiv("2016-12-22", 24.86, 2.48, $sDesc)

$sDesc = "VANGUARD TOTAL BOND MARKET ETF"
FillDiv("2016-02-05", 17.44, 1.74, $sDesc)
; ...
FillDiv("2016-12-07", 16.02, 1.60, $sDesc)
FillDiv("2016-12-29", 1.14, 0, $sDesc)    ; LT Cap Gain
FillDiv("2016-12-29", 2.43, 0, $sDesc)    ; ST Cap Gain

$sDesc = "RDR UNITED COMPANY RUSAL PLC"
FillDiv("2016-11-08", 1230, 0, $sDesc, "Брокер ""YYY"" (Россия)", 832, 643)

Func FillDiv($sDate, $fGross, $fWithheld, $sSecDesc, _
              $sSource = $g_sDefaultIncomeSource, _
              $iCountry = $g_iDefaultCountryCode, _
              $iCurrency = $g_iDefaultCurrencyCode _
)
  ; Проверяю, что строка с описанием источника не пустая.
  If $sSource = "" Then
    MsgBox($MB_SYSTEMMODAL, "Ошибка", "Не задано описание источника выплаты")
    Exit
  EndIf
  ; Проверяю, что доход - положительный.
  If $fGross <= 0 Then
    MsgBox($MB_SYSTEMMODAL, "Ошибка", "Отрицательный или нулевой доход: " & String($fGross))
    Exit
  EndIf
  ; Проверяю, что удержанная сумма - неотрицательная.
  If $fWithheld < 0 Then
    MsgBox($MB_SYSTEMMODAL, "Ошибка", "Отрицательная удержанная сумма: " & String($fWithheld))
    Exit
  EndIf
  ; Проверяю, что удержано меньше, чем поступило.
  If $fWithheld >= $fGross Then
    MsgBox($MB_SYSTEMMODAL, "Ошибка", "Удержано больше(равно), чем поступило: " & String($fWithheld))
    Exit
  EndIf
  ; Проверяю, что дата - существующая.
  If Not _DateIsValid($sDate) Then
    MsgBox($MB_SYSTEMMODAL, "Ошибка", "Несуществующая дата: " & $sDate)
    Exit
  EndIf
  ; Перевожу строку даты в числа.
  Local $aDate, $aTime
  _DateTimeSplit($sDate, $aDate, $aTime)
  Local $iYear = $aDate[1], $iMonth = $aDate[2], $iDay = $aDate[3]
  ; Проверяю, что год введённой даты равен году налогового периода.
  If $iYear <> $g_iYear Then
    MsgBox($MB_SYSTEMMODAL, "Ошибка", "Неверный год в дате: " & $sDate)
    Exit
  Endif
  ; Проверяю, что дата - рабочий день.
  if _DateToDayOfWeekISO($iYear, $iMonth, $iDay) > 5 Then
    MsgBox($MB_SYSTEMMODAL, "Ошибка", "Дата - выходной: " & $sDate)
    Exit
  EndIf
  ; Нажимаю "Добавить источник выплат".
  MouseClick($MOUSE_CLICK_LEFT, $g_ixAddSrc, $g_iYaddSrc)
  ; Жду появления окна "Доход".
  Local $sTitleIncome = "Доход"
  WinWaitActive($sTitleIncome)
  ; Заполняю "Наименование источника выплаты"
  Local $sIncomeSrc = StringFormat("%s Дивиденд от %02d.%02d.%d %s", $sSource, $iDay, $iMonth, $iYear, $sSecDesc)
  ControlSetText($sTitleIncome, "", "[CLASS:TMaskedEdit; INSTANCE:2]", $sIncomeSrc)
  ; Заполняю "Код страны".
  OpenListViewThenFindAndSelectNumber($sTitleIncome, "[CLASS:TButton; INSTANCE:1]", "Справочник стран", $iCountry)
  If @error Then
    MsgBox($MB_SYSTEMMODAL, "Ошибка", "Неизвестный код страны: " & String($iCountry))
    Exit
  EndIf
  ; Нажимаю "Да".
  ControlClick($sTitleIncome, "", "[CLASS:TButton; INSTANCE:2]")
  ; Жду появления основного окна.
  WinWaitActive($g_sTitleMain);
  ; Заполняю "Дата получения дохода".
  SetDateTimePicker($g_sTitleMain, "[CLASS:TDateTimePicker; INSTANCE:2]", $iDay, $iMonth, $iYear)
  ; Заполняю "Дата уплаты налога".
  SetDateTimePicker($g_sTitleMain, "[CLASS:TDateTimePicker; INSTANCE:1]", $iDay, $iMonth, $iYear)
  ; Заполняю "Код и наименование валюты".
  OpenListViewThenFindAndSelectNumber($g_sTitleMain, "[CLASS:TButton; INSTANCE:3]", "Справочник валют", $iCurrency)
  If @error Then
    MsgBox($MB_SYSTEMMODAL, "Ошибка", "Неизвестный код валюты: " & String($iCurrency))
    Exit
  EndIf
  ; Отмечаю "Автоматическое определение курса валюты".
  ControlClick($g_sTitleMain, "", "[CLASS:TCheckBox; INSTANCE:1]")
  ; Заполняю "Код дохода".
  OpenListViewThenFindAndSelectNumber($g_sTitleMain, "[CLASS:TButton; INSTANCE:2]", "Справочник видов доходов", $g_iIncomeTypeDividends)
  If @error Then
    MsgBox($MB_SYSTEMMODAL, "Ошибка", "Неизвестный код дохода: " & String($g_iIncomeTypeDividends))
    Exit
  EndIf
  ; Заполняю "Полученный доход" - "В иностранной валюте"
  ControlSetText($g_sTitleMain, "", "[CLASS:TMaskedEdit; INSTANCE:3]", FloatToStringRus($fGross))
  ; Заполняю "Налог, уплаченный в иностранном государстве" - "В иностранной валюте".
  ControlSetText($g_sTitleMain, "", "[CLASS:TMaskedEdit; INSTANCE:2]", FloatToStringRus($fWithheld))
EndFunc

Func OpenListViewThenFindAndSelectNumber($sTitleParent, $sCtrlButtonOpen, $sTitleLv, $iNumberToSelect)
  ; Нажимаю кнопку вызова списка.
  ControlClick($sTitleParent, "", $sCtrlButtonOpen)
  ; Жду появления окна списка.
  WinWaitActive($sTitleLv)
  ; Ищу строку в списке.
  Local $sCtrlLv = "[CLASS:TListView; INSTANCE:1]"
  Local $iIndex = ControlListView($sTitleLv, "", $sCtrlLv, "FindItem" , String($iNumberToSelect))
  if $iIndex = -1 Then
    SetError(1)
    Return
  Endif
  ; Выбираю строку в списке.
  ControlListView($sTitleLv, "", $sCtrlLv, "Select" , $iIndex)
  ; Нажимаю кнопку подтвержения выбора.
  Local $sCtrlButtonConfirm = "[CLASS:TButton; INSTANCE:2]"
  ControlClick($sTitleLv, "", $sCtrlButtonConfirm)
  ; Жду появления родительского окна.
  WinWaitActive($sTitleParent)
EndFunc

Func SetDateTimePicker($sTitle, $sCtrl, $iDay, $iMonth, $iYear)
  ; Устанавливаю фокус на элемент выбора даты.
  ControlFocus($sTitle, "", $sCtrl)
  Local $iDelayMs = 100
  ; Если дата в элементе вводится не первый раз, то фокус будет на годе.
  ; Нажимаю два раза стрелку влево, чтобы установить фокус на день.
  Local Static $iCalls = 0
  $iCalls += 1
  If $iCalls > 2 Then
    ControlSend($sTitle, "", $sCtrl, "{LEFT}")
    Sleep($iDelayMs)
    ControlSend($sTitle, "", $sCtrl, "{LEFT}")
    Sleep($iDelayMs)
  Endif
  ; Ввожу компонент даты, нажимаю стрелку вправо.
  ControlSend($sTitle, "", $sCtrl, String($iDay))
  Sleep($iDelayMs)
  ControlSend($sTitle, "", $sCtrl, "{RIGHT}")
  Sleep($iDelayMs)
  ControlSend($sTitle, "", $sCtrl, String($iMonth))
  Sleep($iDelayMs)
  ControlSend($sTitle, "", $sCtrl, "{RIGHT}")
  Sleep($iDelayMs)
  ControlSend($sTitle, "", $sCtrl, String($iYear))
  Sleep($iDelayMs)
EndFunc

Func OpenDecl()
  ; Проверяю, что разрешение экрана, не менее 1024 в ширину
  If @DeskTopWidth < 1024 Then
    MsgBox($MB_SYSTEMMODAL, "Ошибка", "Разрешение экрана не поддерживается: " & @DeskTopWidth & "x" & @DeskTopHeight)
    Exit
  EndIf
  ; Проверяю, что запущен только один экземляр этого скрипта.
  If _Singleton(@ScriptName, 1) = 0 Then
    MsgBox($MB_SYSTEMMODAL, "Ошибка", "Скрипт " & @ScriptName & " уже запущен")
    Exit
  EndIf
  ; Проверяю, что Декларация не запущена.
  If ProcessExists($g_sDeclBinary) Then
    MsgBox($MB_SYSTEMMODAL, "Ошибка", "Программа Декларация уже запущена")
    Exit
  Endif
  ; Запускаю программу.
  Run($g_sPath)
  ; Жду появления основного окна.
  WinWaitActive($g_sTitleMain)
  ; Разворачиваю окно на весь экран.
  WinSetState($g_sTitleMain, "", @SW_MAXIMIZE)
  ; Отмечаю "Имеются доходы - В иностранной валюте".
  ControlClick($g_sTitleMain, "", "[TEXT:В иностранной валюте]")
  ; Перехожу в раздел "Доходы за пределами РФ".
  MouseClick($MOUSE_CLICK_LEFT, $g_iXxRus, $g_iYxRus)
  ; Жду появления окна с сообщением "Не введен номер инспекции".
  Local $sTitleErr = "Проверка: Задание условий"
  WinWaitActive($sTitleErr)
  ; Нажимаю "Пропустить".
  ControlClick($sTitleErr, "", "[TEXT:Пропустить]")
  ; Жду появления основного окна.
  WinWaitActive($g_sTitleMain)
EndFunc

Func FloatToStringRus($fVal)
  ; Преобразую число в строку, заменяю точку на запятую.
  Return StringReplace(String($fVal), ".", ",")
EndFunc