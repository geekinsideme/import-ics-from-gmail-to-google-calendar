Import .ics file from Gmail attachment to Google Calendar (Google Apps Script)

Gmail宛に添付された.icsファイルの内容をGoogleカレンダーに取り込む Google Apps Script

- なんらかの方法で、.ics ファイルをGmailの自分宛てに送付する
- このスクリプトを実行すると、その.icsファイルに登録されているイベントが、指定したGooleカレンダーにインポートされる。
- .icsファイルの取り込みを終えると、当該のメールはゴミ箱に入る
- このスクリプトを定期的に実行させることで、所望するカレンダーとGoogleカレンダーの擬似的な同期が行われる。
- (Googleカレンダーにはinternet上の.icsを購読する機能が備わっているが、1日に一回程度しか同期せず、使い物にならない)
