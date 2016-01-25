Import .ics file from Gmail attachment to Google Calendar automatically using Google Apps Script programmatically

Gmail宛に添付された.icsファイルの内容をGoogleカレンダーに取り込む Google Apps Script

- なんらかの方法で、.ics ファイルをGmailの自分宛てに送付する
- このスクリプトを実行すると、その.icsファイルに登録されているイベントが、指定したGooleカレンダーにインポートされる。
- .icsファイルの取り込みを終えると、当該のメールはゴミ箱に入る
- このスクリプトを定期的に実行させる(Google Apps Scriptの機能)ことで、所望するカレンダーとGoogleカレンダーの擬似的な同期が行われる。
- (Googleカレンダーにはinternet上の.icsを購読する機能が備わっているが、1日に一回程度しか同期せず、使い物にならない)

- 注意:このスクリプトでは、指定されたGoogleカレンダーの内容(予定)をいったんすべて削除してから、.icsファイルの内容を登録する。よって、Googleカレンダーで新しい予定表を作ってから指定しなければ、悲しいことになる。
