const ExcelJS = require('exceljs');
const path = require('path');

async function main() {
  const wb = new ExcelJS.Workbook();

  const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };
  const editFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFDE7' } };
  const headerFont = { bold: true, size: 11 };
  const normalFont = { size: 10 };
  const redFont = { color: { argb: 'FFFF0000' }, size: 10 };
  const wrapAlign = { wrapText: true, vertical: 'top' };
  const thinBorder = {
    top: { style: 'thin' }, bottom: { style: 'thin' },
    left: { style: 'thin' }, right: { style: 'thin' }
  };
  const colWidths = [18, 28, 55, 55, 32];
  const headers = ['セクション', '項目', '現状の文言', '変更後（記入欄）', '備考'];

  function setupSheet(ws, data, warnings = {}) {
    // Headers
    headers.forEach((h, i) => {
      const cell = ws.getCell(1, i + 1);
      cell.value = h;
      cell.font = headerFont;
      cell.fill = headerFill;
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = thinBorder;
    });
    // Column widths
    colWidths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
    // Data
    data.forEach((row, ri) => {
      const rowIdx = ri + 2;
      row.forEach((val, ci) => {
        const cell = ws.getCell(rowIdx, ci + 1);
        cell.value = val || '';
        cell.alignment = wrapAlign;
        cell.border = thinBorder;
        cell.font = normalFont;
        if (ci === 3) cell.fill = editFill;
      });
      if (warnings[rowIdx]) {
        ws.getCell(rowIdx, 5).font = redFont;
      }
    });
    ws.views = [{ state: 'frozen', ySplit: 1 }];
  }

  // Sheet 1: Pattern A
  const ws1 = wb.addWorksheet('Pattern A（白）');
  const dataA = [
    ['Hero','バッジ','🎬 TikTokクリエイター向け','',''],
    ['Hero','見出し','好きな商品を紹介して報酬を得よう','',''],
    ['Hero','サブテキスト','TCSは、TikTokクリエイターと人気ブランドをつなぐアフィリエイトプログラム。専用ポータルで案件管理、特別料率で高報酬。','',''],
    ['Hero','CTAボタン','LINEで無料登録','',''],
    ['Hero','注釈','最短30秒で登録完了 ・ 費用は一切かかりません','',''],
    ['実績','ラベル','Results','',''],
    ['実績','タイトル','実績で選ばれています','',''],
    ['実績','数字1','120+ / 提携クリエイター','',''],
    ['実績','数字2','40+ / 取扱ブランド・商品','',''],
    ['実績','数字3','20% / 最大報酬料率','',''],
    ['実績','数字4','¥59万 / 月間TOP報酬額','',''],
    ['選ばれる理由','タイトル','TCSが選ばれる理由','',''],
    ['選ばれる理由','サブ','一般的なアフィリエイトとは違う、クリエイターファーストの仕組み。','',''],
    ['選ばれる理由','特徴1','💰 特別料率で高報酬 / TikTok Shopの公開料率よりも高い特別料率を提供。','',''],
    ['選ばれる理由','特徴2','📱 専用ポータルで案件管理 / 案件の検索・応募・進捗管理がすべてポータルで完結。','',''],
    ['選ばれる理由','特徴3','🎁 サンプル無料提供 / 紹介する商品のサンプルを無料でお届け。','',''],
    ['選ばれる理由','特徴4','⚡ 審査なしで即参加OK / 一部の案件は審査不要で即参加可能。','',''],
    ['選ばれる理由','特徴5','🎬 動画もLIVEもOK / 投稿スタイルは自由。','',''],
    ['選ばれる理由','特徴6','🤝 専任サポート付き / LINEですぐに相談可能。','',''],
    ['ポータル','タイトル','専用ポータルで案件を一括管理','',''],
    ['ポータル','説明','登録するとクリエイター専用ポータルにアクセスできます。','',''],
    ['ポータル','商品例1','スキンケアセット A / 報酬率 15%','',''],
    ['ポータル','商品例2','オーガニックコーヒー / 報酬率 12%','',''],
    ['ポータル','商品例3','リップグロス 新色 / 報酬率 20%','',''],
    ['始め方','タイトル','始めるのはとても簡単','',''],
    ['始め方','Step1','LINEで無料登録 / 30秒で登録完了','',''],
    ['始め方','Step2','案件を選んで応募','',''],
    ['始め方','Step3','商品を紹介 / サンプルが届いたら自由に紹介','',''],
    ['始め方','Step4','報酬GET / 売れた分だけ報酬発生','',''],
    ['クリエイターの声','タイトル','クリエイターの声','',''],
    ['クリエイターの声','声1','「ポータルが使いやすくて案件選びが楽」/ 美容系 3.2万人','','※プレースホルダー。実際の声に差し替え推奨'],
    ['クリエイターの声','声2','「料率が高い。月20万以上稼げてます」/ ライフスタイル系 5.8万人','','※プレースホルダー'],
    ['クリエイターの声','声3','「サンプルが届くので良い商品だけ紹介できる」/ フード系 1.5万人','','※プレースホルダー'],
    ['FAQ','Q1','費用はかかりますか？ → 完全無料','',''],
    ['FAQ','Q2','フォロワー数の条件は？ → 特に下限なし','',''],
    ['FAQ','Q3','対象ジャンルは？ → 幅広いジャンル','',''],
    ['FAQ','Q4','報酬の受け取りは？ → TikTok Shop経由','',''],
    ['FAQ','Q5','LIVE配信のみ？ → 動画投稿でもOK','',''],
    ['CTA','見出し','あなたのTikTokで報酬を得よう','',''],
    ['CTA','CTAボタン','LINEで無料登録する','',''],
  ];
  setupSheet(ws1, dataA, { 32: true, 33: true, 34: true });

  // Sheet 2: Pattern B
  const ws2 = wb.addWorksheet('Pattern B（ダーク）');
  const dataB = [
    ['Hero','アイキャッチ','TikTok Creator Partner Program','',''],
    ['Hero','見出し','TikTokで稼げるを、本気で。','',''],
    ['Hero','サブ','人気ブランド × 特別料率 × 専用ツール','',''],
    ['Hero','数字1','¥59万 / 月間TOP報酬','',''],
    ['Hero','数字2','20% / 最大報酬率','',''],
    ['Hero','数字3','40+ / 取扱商品','',''],
    ['Hero','CTAボタン','今すぐ参加する','',''],
    ['Hero','注釈','30秒で登録完了 ・ 完全無料','',''],
    ['比較','タイトル','普通のアフィリエイトと何が違う？','',''],
    ['比較','一般アフィリ','公開料率のみ/案件自分で探す/サポートなし/サンプル自腹/管理手動','',''],
    ['比較','TCSパートナー','特別料率/ポータルで案件提案/LINE専任サポート/サンプル無料/自動管理','',''],
    ['報酬','タイトル','報酬率が全然違う','',''],
    ['報酬','バー1','5% / 公開料率','',''],
    ['報酬','バー2','15% / TCS特別料率','',''],
    ['報酬','バー3','20% / 招待限定','',''],
    ['始め方','タイトル','4ステップで始められる','',''],
    ['始め方','Step1','LINE登録 / 30秒で完了','',''],
    ['始め方','Step2','案件を選ぶ','',''],
    ['始め方','Step3','紹介する / 動画でもLIVEでもOK','',''],
    ['始め方','Step4','報酬GET','',''],
    ['ポータル','タイトル','クリエイター専用ポータル','',''],
    ['ポータル','機能','案件検索/進捗管理/LINE通知/お気に入り','',''],
    ['ブランド','タイトル','人気ブランドと提携','',''],
    ['ブランド','一覧','こうじや/のむシリカ/UCC/ROSABLU/Vioteras/KINS/リクセル/イングリウッド','','※実ブランド名。公開前に許可確認必要'],
    ['CTA','見出し','始めるなら、今。','',''],
    ['CTA','CTAボタン','今すぐ参加する','',''],
  ];
  setupSheet(ws2, dataB, { 25: true });

  // Sheet 3: 共通メモ
  const ws3 = wb.addWorksheet('共通メモ');
  const memo = [
    ['項目','内容','ステータス'],
    ['LINE登録URL','https://lin.ee/XXXXX','→ 実URLに差替え必要'],
    ['Voices（クリエイターの声）','プレースホルダー','→ 実際の声に差替え推奨'],
    ['Brands（ブランド名）','実名使用','→ 許可確認必要'],
    ['Preview URL - Pattern A','https://rkrymd.github.io/tcs-portal/lp_a.html',''],
    ['Preview URL - Pattern B','https://rkrymd.github.io/tcs-portal/lp_b.html',''],
  ];
  memo.forEach((row, ri) => {
    row.forEach((val, ci) => {
      const cell = ws3.getCell(ri + 1, ci + 1);
      cell.value = val;
      cell.alignment = wrapAlign;
      cell.border = thinBorder;
      if (ri === 0) { cell.font = headerFont; cell.fill = headerFill; }
      else if (ci === 2 && val.startsWith('→')) cell.font = redFont;
      else cell.font = normalFont;
    });
  });
  ws3.getColumn(1).width = 30;
  ws3.getColumn(2).width = 55;
  ws3.getColumn(3).width = 30;
  ws3.views = [{ state: 'frozen', ySplit: 1 }];

  const outputPath = path.join(__dirname, 'LP文言チェックシート.xlsx');
  await wb.xlsx.writeFile(outputPath);
  console.log('Created:', outputPath);
}

main().catch(console.error);
