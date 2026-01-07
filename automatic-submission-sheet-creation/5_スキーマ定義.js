/**
 * OpenAI構造化出力用のJSONスキーマ定義
 * response_format で使用
 */

/**
 * mappingObjから全フィールド名を収集してスキーマプロパティを生成
 */
function generateFieldsSchema() {
  const allFields = new Set();

  // mappingObjから全てのキーを収集
  Object.keys(mappingObj).forEach(mappingKey => {
    mappingObj[mappingKey].forEach(item => {
      allFields.add(item.key);
    });
  });

  // スキーマのpropertiesオブジェクトを生成
  const properties = {};
  const requiredFields = [];

  allFields.forEach(fieldName => {
    properties[fieldName] = {
      type: ["string", "null"],
      description: `${fieldName}の値。不明な場合はnull`
    };
    requiredFields.push(fieldName);
  });

  return {
    type: "object",
    properties: properties,
    required: requiredFields,
    additionalProperties: false
  };
}

/**
 * OpenAI構造化出力用のJSONスキーマ
 */
function getResponseSchema() {
  return {
    name: "media_plan_extraction",
    strict: true,
    schema: {
      type: "object",
      properties: {
        medias: {
          type: "array",
          description: "抽出された媒体情報の配列",
          items: {
            type: "object",
            properties: {
              mediaId: {
                type: "string",
                description: "媒体ID。MEDIA_CONFIGのキーと完全一致させる。例: 'Googleリスティング', 'Yahoo！リスティング', 'GDN', 'YDA', 'Meta', 'LINE', 'X', 'TikTok'"
              },
              confidence: {
                type: ["number", "null"],
                description: "抽出の確信度 (0.0～1.0)。判断が難しい場合は低い値を設定"
              },
              fields: generateFieldsSchema()
            },
            required: ["mediaId", "confidence", "fields"],
            additionalProperties: false
          }
        },
        unmapped_notes: {
          type: ["string", "null"],
          description: "抽出できなかった情報や特記事項。媒体に該当しない情報がある場合に記載"
        }
      },
      required: ["medias", "unmapped_notes"],
      additionalProperties: false
    }
  };
}

/**
 * 利用可能な媒体IDのリストを取得
 */
function getAvailableMediaIds() {
  return Object.keys(MEDIA_CONFIG);
}

/**
 * フロントエンド用: 媒体リストを取得
 */
function getMediaList() {
  return getAvailableMediaIds();
}

/**
 * スキーマ検証: レスポンスが期待形式に準拠しているか確認
 */
function validateResponseSchema(response) {
  const errors = [];

  // mediasが配列か
  if (!Array.isArray(response.medias)) {
    errors.push("medias is not an array");
    return { valid: false, errors };
  }

  // 各媒体の検証
  response.medias.forEach((media, index) => {
    if (!media.mediaId) {
      errors.push(`medias[${index}]: mediaId is missing`);
    }

    if (!media.fields || typeof media.fields !== 'object') {
      errors.push(`medias[${index}]: fields is missing or not an object`);
    }

    // mediaIdがMEDIA_CONFIGに存在するか
    if (media.mediaId && !MEDIA_CONFIG[media.mediaId]) {
      errors.push(`medias[${index}]: mediaId '${media.mediaId}' not found in MEDIA_CONFIG`);
    }
  });

  return {
    valid: errors.length === 0,
    errors
  };
}
