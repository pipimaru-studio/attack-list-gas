// 文字列の整形に使う共通関数 

// ----------------【会社名変換】----------------
// 会社名変換
function convert_companyName(text) {
  // 会社名変換処理（カタカナ全角変換や略称の変換など）
  text = convertKatakana(text)
    .replace(/[\u200B-\u200D\uFEFF\u00A0]/g, '')// 不可視文字を削除
    .replace(/･/g, '・');  // まず半角中黒を全角中黒に統一
  text = toHalfWidth(text) // 英数字・記号をすべて半角に統一
    .replace(/　/g, '') // 全角スペースを削除
    .replace(/\s*([&･・‧])\s*/g, '$1')  // 記号の前後のスペース削除
    .replace(/[（(]\s*株\s*[)）]/g, '株式会社')
    .replace(/[（(]\s*合\s*[)）]/g, '合同会社')
    .replace(/[（(]\s*有\s*[)）]/g, '有限会社')
    .replace(/(株式会社|合同会社|有限会社)\s+/g, '$1')  // 株式（合同、有限）会社　右側のスペースを削除
    .replace(/\s+(?=株式会社|合同会社|有限会社)/g, '');  // 株式（合同、有限）会社　左側のスペースを削除
  return text.trim();
}

// カタカナ全角変換(https://hi3103.net/notes/google/1394)
function convertKatakana(text) {
  const zenkata = [
  'ア', 'イ', 'ウ', 'エ', 'オ',
  'カ', 'キ', 'ク', 'ケ', 'コ',
  'サ', 'シ', 'ス', 'セ', 'ソ',
  'タ', 'チ', 'ツ', 'テ', 'ト',
  'ナ', 'ニ', 'ヌ', 'ネ', 'ノ',
  'ハ', 'ヒ', 'フ', 'ヘ', 'ホ',
  'マ', 'ミ', 'ム', 'メ', 'モ',
  'ヤ', 'ユ', 'ヨ',
  'ラ', 'リ', 'ル', 'レ', 'ロ',
  'ワ', 'ヲ', 'ン',
  'ガ', 'ギ', 'グ', 'ゲ', 'ゴ',
  'ザ', 'ジ', 'ズ', 'ゼ', 'ゾ',
  'ダ', 'ヂ', 'ヅ', 'デ', 'ド',
  'バ', 'ビ', 'ブ', 'ベ', 'ボ',
  'パ', 'ピ', 'プ', 'ペ', 'ポ',
  'ァ', 'ィ', 'ゥ', 'ェ', 'ォ',
  'ャ', 'ュ', 'ョ',
  'ッ', 'ヴ',
  '。', '、'];

  //半角カタカナ一覧
  const hankata = [
  'ｱ', 'ｲ', 'ｳ', 'ｴ', 'ｵ',
  'ｶ', 'ｷ', 'ｸ', 'ｹ', 'ｺ',
  'ｻ', 'ｼ', 'ｽ', 'ｾ', 'ｿ',
  'ﾀ', 'ﾁ', 'ﾂ', 'ﾃ', 'ﾄ',
  'ﾅ', 'ﾆ', 'ﾇ', 'ﾈ', 'ﾉ',
  'ﾊ', 'ﾋ', 'ﾌ', 'ﾍ', 'ﾎ',
  'ﾏ', 'ﾐ', 'ﾑ', 'ﾒ', 'ﾓ',
  'ﾔ', 'ﾕ', 'ﾖ',
  'ﾗ', 'ﾘ', 'ﾙ', 'ﾚ', 'ﾛ',
  'ﾜ', 'ｦ', 'ﾝ',
  'ｶﾞ', 'ｷﾞ', 'ｸﾞ', 'ｹﾞ', 'ｺﾞ',
  'ｻﾞ', 'ｼﾞ', 'ｽﾞ', 'ｾﾞ', 'ｿﾞ',
  'ﾀﾞ', 'ﾁﾞ', 'ﾂﾞ', 'ﾃﾞ', 'ﾄﾞ',
  'ﾊﾞ', 'ﾋﾞ', 'ﾌﾞ', 'ﾍﾞ', 'ﾎﾞ',
  'ﾊﾟ', 'ﾋﾟ', 'ﾌﾟ', 'ﾍﾟ', 'ﾎﾟ',
  'ｧ', 'ｨ', 'ｩ', 'ｪ', 'ｫ',
  'ｬ', 'ｭ', 'ｮ',
  'ｯ', 'ｳﾞ',
  '｡', '､'];

  //最終的に返す変数を定義
  let result = '';

  //変換する対象文字列をセット
  const input = hankata;
  const output = zenkata;

  //引数に値が格納されているかチェック
  if(typeof text === 'undefined') {
    // result = 'エラー：全角カタカナに変換する文字列が指定されていません。';
    // エラーの場合、そのまま返す
    result = text;
  }else{
    //引数で渡された文字列を定数にセット
    const textStr = text;

    //1文字ずつ分割して配列に格納
    const textArr = textStr.split('');

    //文字数が0でなければ実行
    if(textArr.length !== 0){
      //文字を再格納する配列を定義
      let array = [];

      //分割した文字の数だけループを回し、もし濁点・半濁点だった場合は1つ前の配列の中身とセットにして array に格納
      for (let i = 0; i < textArr.length; i++) {
        if (textArr[i] == 'ﾞ' || textArr[i] == 'ﾟ') {
          array[array.length - 1] = (textArr[i - 1] + textArr[i]);
        } else {
          array.push(textArr[i]);
        }
      }

      //再格納した文字の数だけループを回し、もし半角カナがあったら全角カナに直して result へ格納
      for (let j = 0; j < array.length; j++) {
        let index = input.indexOf(array[j]);
        if (index == -1) {
          result = result + array[j];
        } else {
          result = result + output[index];
        }
      }
    }
  }

  //結果を返す
  return result;
} 


// ----------------【電話番号変換】----------------
function convert_phonenumber(text) {
  // 英数字半角変換し入力ルールにフォーマットを整える
  text = toHalfWidth(text) // 英数字・記号をすべて半角に統一
    .replace(/[‐−—―]/g, '-') // 全角ハイフン・マイナス記号を半角ハイフンに変換                              
    .replace(/\(/g, '-') // (をハイフンに変換
    .replace(/\)/g, '-') // )をハイフンに変換
    .replace(/[^0-9-]/g, "") // 数字とハイフン以外を空欄に置換
    .replace(/^-+/, '') // 文字列の先頭にある "-" を "" に置換
    .replace(/-+$/g, "")// 末尾のハイフンをすべて削除
    .replace(/\s+/g, "")// 不要なスペースを削除

   text = text.trim();
    
  // ハイフンが1つ以上含まれていたら、それをそのまま使う
  if (text.includes('-')) {
    return text;
  }
  // ハイフンがないなら、ハイフンを挿入する
  return formatPhoneNumberWithHyphen(text);                                              
}

// 電話番号整形：ハイフン挿入（例：03-1234-5678）
function formatPhoneNumberWithHyphen(number) {
  number = number.replace(/[^0-9]/g, ""); // 数字だけに

  if (number.length === 10) {
    if (number.startsWith("03") || number.startsWith("06")) {
      // 東京・大阪 → 2-4-4
      return number.replace(/(\d{2})(\d{4})(\d{4})/, "$1-$2-$3");
    } else {
      // その他の固定電話（名古屋・福岡など）→ 3-3-4
      return number.replace(/(\d{3})(\d{3})(\d{4})/, "$1-$2-$3");
    }
  } else if (number.length === 11) {
    // 携帯番号など → 3-4-4
    return number.replace(/(\d{3})(\d{4})(\d{4})/, "$1-$2-$3");
  } else {
    return number; // ハイフン整形不可（そのまま返す）
  }
}

// ----------------【共通処理】----------------
// 英数字・記号をすべて半角に統一
function toHalfWidth(text) {
  return text.replace(/[Ａ-Ｚａ-ｚ０-９！-～]/g, s => String.fromCharCode(s.charCodeAt(0) - 65248));
}

// ----------------【email変換】----------------
function convert_email(text) {
  // 受け取ったデータが空っぽ（null や undefined）だったら、何もせずに空の文字列 "" を返す
  if (!text) {
    return "";
  }
 
 // 英数字・記号すべて半角に
text = text.replace(/[\uFF01-\uFF5E]/g, function(s) {
  return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
});

// Logger.log ("半角変換後：" + text );
// Logger.log("コロンのUnicode：" + text.charCodeAt(5));
// for (let i = 0; i < text.length; i++) {
//   Logger.log(`text[${i}] = '${text[i]}' / Unicode = ${text.charCodeAt(i)}`);
// }

//改行や空白を除去（正規表現に正しくかけるため）
text = text.trim();

  // 文字列からアドレス部分のみ抽出
 const emailMatch = text.match(/^([Ee]mail|[Ee]-[mM]ail|[mM]ail)\s*[:]?\s*(.+)$/i);

  //   `emailMatch` が見つかっていて（`null` でない）、
  //   さらにその配列の3番目の要素（メールアドレス）が存在する場合に、アドレス部分のみ抽出。
if (emailMatch && emailMatch[2]) {
  // Logger.log("match[0] > " + emailMatch[0]);
  // Logger.log("match[1] > " + emailMatch[1]);
  // Logger.log("match[2] > " + emailMatch[2]);
  text = emailMatch[2]
} 
else {
  // Logger.log("メールアドレス形式にマッチしませんでした：" + text);
}

  // 日本語を空欄に置換
  text = text.replace(/[\u3000-\u303f\u3040-\u309f\u30a0-\u30ff\uff00-\uff9f\u4e00-\u9faf\u3400-\u4dbf]/g, '');

  // その他の不要な空白文字を削除
  text = text.replace(/\s/g, '');

  return text.trim();
}


