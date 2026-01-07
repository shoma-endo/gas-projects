/**
 * JSONレスポンスをクリーンアップ
 *
 * 【なぜこの関数が必要か】
 * - OpenAI APIが時折、過剰な改行や空白を含むJSONを返すことがある
 * - 特に構造化出力で大量の null フィールドがある場合、数千行に膨れ上がることがある
 * - これによりJSONパースエラーやログ容量オーバーが発生する
 *
 * @param {string} jsonString - クリーンアップするJSON文字列
 * @returns {string} クリーンアップされたJSON文字列
 */
function cleanJsonResponse(jsonString) {
  try {
    // 【なぜ】複数の連続する改行を1つに統合
    // → AIが "field": null,\n\n\n\n のような出力をすることがある
    let cleaned = jsonString.replace(/\n\s*\n\s*\n+/g, '\n');

    // 【なぜ】文字列値内の制御文字をエスケープ
    // → 文字列値内の改行やタブがJSONパースエラーを起こす場合があるため
    // → ただし、既にエスケープされている場合は重複しないように注意
    cleaned = cleaned.replace(/"([^"\\]*(\\.[^"\\]*)*)"/g, function(match, p1) {
      // 文字列値内の未エスケープの改行・タブを修正
      let fixed = p1
        .replace(/\n/g, '\\n')
        .replace(/\r/g, '\\r')
        .replace(/\t/g, '\\t');
      return '"' + fixed + '"';
    });

    // 【なぜ】括弧の前後の過剰な空白を削除
    // → { \n\n\n "field" のようなフォーマットを { "field" に正規化
    cleaned = cleaned
      .replace(/{\s+/g, '{')      // { の後の空白
      .replace(/\s+}/g, '}')      // } の前の空白
      .replace(/\[\s+/g, '[')     // [ の後の空白
      .replace(/\s+]/g, ']')      // ] の前の空白
      .replace(/,\s+/g, ',')      // , の後の空白（改行は保持）
      .replace(/:\s+/g, ':');     // : の後の空白

    // 【なぜ】全体をトリム
    // → 先頭・末尾の不要な空白を除去
    cleaned = cleaned.trim();

    return cleaned;
  } catch (e) {
    // クリーンアップ自体が失敗した場合は元の文字列を返す
    Logger.log(`⚠️ JSONクリーンアップ失敗: ${e.message}`);
    return jsonString;
  }
}

/**
 * ワンショットで与件全文から全媒体情報を抽出
 * OpenAI構造化出力（response_format）を使用
 *
 * @param {string} requestText - 与件全文
 * @param {string} submissionId - リクエストID（ロギング用）
 * @returns {Object} { medias: [...], unmapped_notes: "..." }
 */
function analyzeTextAI(requestText, submissionId) {

  const startTime = Date.now();

  // ====== プロンプト生成 ======
  const systemPrompt = promptObj.getSystemPrompt();
  const userPrompt = promptObj.getUserPrompt(requestText);

  // ====== OpenAI構造化出力呼び出し ======
  const apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!apiKey) {
    throw new Error("❌ OPENAI_API_KEYが設定されていません");
  }

  const url = "https://api.openai.com/v1/chat/completions";
  const schema = getResponseSchema();

  const payload = {
    model: CONFIG.AI_MODEL,
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userPrompt }
    ],
    temperature: CONFIG.AI_TEMPERATURE,
    max_tokens: CONFIG.AI_MAX_TOKENS,
    response_format: {
      type: "json_schema",
      json_schema: schema
    }
  };

  const options = {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${apiKey}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  // リトライロジック付きでAPI呼び出し
  let response;
  let attempt = 0;
  const maxRetries = CONFIG.AI_MAX_RETRIES;

  while (attempt < maxRetries) {
    try {
      response = UrlFetchApp.fetch(url, options);
      const statusCode = response.getResponseCode();

      if (statusCode === 200) {
        break; // 成功
      } else if (statusCode === 429) {
        // レート制限
        attempt++;
        if (attempt >= maxRetries) {
          throw new Error(`❌ OpenAI API rate limit exceeded after ${maxRetries} retries`);
        }
        const waitTime = Math.pow(2, attempt) * 1000; // 指数バックオフ
        Logger.log(`⏳ Rate limit hit. Waiting ${waitTime}ms before retry ${attempt}/${maxRetries}`);
        Utilities.sleep(waitTime);
      } else {
        // その他のHTTPエラー
        throw new Error(`❌ OpenAI API error ${statusCode}: ${response.getContentText()}`);
      }
    } catch (e) {
      attempt++;
      if (attempt >= maxRetries) {
        throw new Error(`❌ OpenAI API request failed after ${maxRetries} retries: ${e.message}`);
      }
      const waitTime = Math.pow(2, attempt) * 1000;
      Logger.log(`⏳ Request error. Waiting ${waitTime}ms before retry ${attempt}/${maxRetries}: ${e.message}`);
      Utilities.sleep(waitTime);
    }
  }

  // ====== レスポンスパース ======
  let data;
  try {
    data = JSON.parse(response.getContentText());
  } catch (e) {
    throw new Error(`❌ Failed to parse OpenAI response: ${e.message}\n${response.getContentText()}`);
  }

  if (!data?.choices?.[0]?.message?.content) {
    throw new Error(`❌ Invalid OpenAI response structure:\n${JSON.stringify(data, null, 2)}`);
  }

  // 構造化出力の場合、contentは既にJSON文字列
  let result;
  try {
    // 【なぜ】JSONをクリーンアップしてからパース
    // → AIが大量の空白行や不正なフォーマットを返す場合があるため
    const rawContent = data.choices[0].message.content;

    // 【なぜ】異常に大きいレスポンスを検出
    // → 正常なレスポンスは数KB程度、1MB超えは異常
    if (rawContent.length > 1000000) {
      Logger.log(`⚠️ 警告: AIレスポンスが異常に大きい (${rawContent.length}文字)`);
    }

    // 【なぜ】JSONをクリーンアップ
    // → 過剰な改行、空白、未エスケープの制御文字を除去して正常なJSONに整形
    const cleanedContent = cleanJsonResponse(rawContent);

    result = JSON.parse(cleanedContent);
  } catch (e) {
    // 【なぜ】詳細なエラー情報をログ出力
    // → JSONパースエラーの原因を特定しやすくするため
    const rawContent = data.choices[0].message.content;

    // エラー位置周辺のコンテキストを抽出
    let errorContext = '';
    if (e.message && e.message.includes('position')) {
      const match = e.message.match(/position (\d+)/);
      if (match) {
        const pos = parseInt(match[1]);
        const start = Math.max(0, pos - 100);
        const end = Math.min(rawContent.length, pos + 100);
        errorContext = `\n\nエラー位置周辺:\n"${rawContent.substring(start, end).replace(/\n/g, '\\n')}"`;
      }
    }

    // 【なぜ】エラーメッセージを切り詰め
    // → 元のJSONが巨大だとログが溢れるため、最初の500文字だけ表示
    const preview = rawContent.length > 500
      ? rawContent.substring(0, 500) + `... (残り${rawContent.length - 500}文字)`
      : rawContent;

    // 完全なJSONをログに出力（デバッグ用）
    Logger.log(`🔍 [デバッグ] JSONパースエラー詳細:\n${e.message}${errorContext}`);
    Logger.log(`🔍 [デバッグ] 完全なJSON (最初の2000文字):\n${rawContent.substring(0, 2000)}`);

    throw new Error(`❌ Failed to parse structured output: ${e.message}${errorContext}\n\nプレビュー:\n${preview}`);
  }

  // ====== スキーマ検証 ======
  const validation = validateResponseSchema(result);
  if (!validation.valid) {
    Logger.log(`⚠️ Schema validation warnings for ${submissionId}:\n${validation.errors.join('\n')}`);
    // 警告のみでエラーにはしない（unmapped_notesに追記）
    if (result.unmapped_notes) {
      result.unmapped_notes += `\n[検証警告] ${validation.errors.join(', ')}`;
    } else {
      result.unmapped_notes = `[検証警告] ${validation.errors.join(', ')}`;
    }
  }

  // ====== ロギング ======
  const elapsed = Date.now() - startTime;
  const usage = data.usage || {};
  Logger.log(`✅ AI抽出完了 [${submissionId}] ${elapsed}ms | tokens: ${usage.total_tokens || 'N/A'} | medias: ${result.medias.length}`);

  // 媒体ごとのログ
  result.medias.forEach((media, idx) => {
    Logger.log(`  [${idx}] ${media.mediaId} (confidence: ${media.confidence || 'N/A'})`);
  });

  if (result.unmapped_notes) {
    Logger.log(`  [備考] ${result.unmapped_notes}`);
  }

  return result;
}

