use wasm_bindgen::prelude::*;
use serde::{Deserialize, Serialize};

#[wasm_bindgen]
extern "C" {
    #[wasm_bindgen(js_namespace = console)]
    fn log(s: &str);
}

// console_logマクロは未使用のためコメントアウト
// macro_rules! console_log {
//     ($($t:tt)*) => (log(&format_args!($($t)*).to_string()))
// }

#[wasm_bindgen(start)]
pub fn main() {
    console_error_panic_hook::set_once();
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableData {
    pub headers: Vec<String>,
    pub rows: Vec<Vec<String>>,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct CompareOptions {
    pub trim: bool,
    pub case_insensitive: bool,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct CompareInput {
    pub left_headers: Vec<String>,
    pub left_rows: Vec<Vec<String>>,
    pub right_headers: Vec<String>,
    pub right_rows: Vec<Vec<String>>,
    pub key: String,
    pub options: CompareOptions,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct CompareOutput {
    pub result: TableData,
    pub left_only: TableData,
    pub right_only: TableData,
    pub duplicates: TableData,
    pub log: Vec<(String, String)>,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct SplitInput {
    pub headers: Vec<String>,
    pub rows: Vec<Vec<String>>,
    pub key: String,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct SplitPart {
    pub key_value: String,
    pub table: TableData,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct SplitOutput {
    pub parts: Vec<SplitPart>,
}

fn normalize_key(key: &str, options: &CompareOptions) -> String {
    let mut normalized = key.to_string();
    if options.trim {
        normalized = normalized.trim().to_string();
    }
    if options.case_insensitive {
        normalized = normalized.to_lowercase();
    }
    normalized
}

#[wasm_bindgen]
pub fn compare_files(input_json: &str) -> String {
    let input: CompareInput = serde_json::from_str(input_json)
        .expect("Failed to parse CompareInput");
    
    let left_key_idx = input.left_headers.iter()
        .position(|h| h == &input.key)
        .expect("Key column not found in left headers");
    let right_key_idx = input.right_headers.iter()
        .position(|h| h == &input.key)
        .expect("Key column not found in right headers");

    // Normalize keys and build maps
    let mut left_map: std::collections::HashMap<String, Vec<usize>> = std::collections::HashMap::new();
    for (idx, row) in input.left_rows.iter().enumerate() {
        if let Some(key_val) = row.get(left_key_idx) {
            let normalized = normalize_key(key_val, &input.options);
            left_map.entry(normalized).or_insert_with(Vec::new).push(idx);
        }
    }

    let mut right_map: std::collections::HashMap<String, Vec<usize>> = std::collections::HashMap::new();
    for (idx, row) in input.right_rows.iter().enumerate() {
        if let Some(key_val) = row.get(right_key_idx) {
            let normalized = normalize_key(key_val, &input.options);
            right_map.entry(normalized).or_insert_with(Vec::new).push(idx);
        }
    }

    // Build result headers
    let mut result_headers: Vec<String> = input.left_headers.iter()
        .map(|h| format!("L__{}", h))
        .collect();
    result_headers.extend(input.right_headers.iter().map(|h| format!("R__{}", h)));
    result_headers.push("match_status".to_string());
    result_headers.push("diff_cols".to_string());
    result_headers.push("dup_key_flag".to_string());

    let mut result_rows: Vec<Vec<String>> = Vec::new();
    let mut left_only_rows: Vec<Vec<String>> = Vec::new();
    let mut right_only_rows: Vec<Vec<String>> = Vec::new();
    let mut duplicates_rows: Vec<Vec<String>> = Vec::new();

    let mut processed_keys: std::collections::HashSet<String> = std::collections::HashSet::new();

    // Find duplicates first
    for (normalized_key, left_indices) in &left_map {
        if left_indices.len() > 1 {
            for &idx in left_indices {
                let left_row = &input.left_rows[idx];
                let mut full_row: Vec<String> = left_row.clone();
                full_row.resize(result_headers.len() - 3, String::new());
                full_row.push("left_only".to_string());
                full_row.push(String::new());
                full_row.push("1".to_string());
                duplicates_rows.push(full_row);
            }
            processed_keys.insert(normalized_key.clone());
        }
    }
    for (normalized_key, right_indices) in &right_map {
        if right_indices.len() > 1 && !processed_keys.contains(normalized_key) {
            for &idx in right_indices {
                let right_row = &input.right_rows[idx];
                let mut full_row: Vec<String> = Vec::new();
                full_row.resize(input.left_headers.len(), String::new());
                full_row.extend(right_row.clone());
                full_row.resize(result_headers.len() - 3, String::new());
                full_row.push("right_only".to_string());
                full_row.push(String::new());
                full_row.push("1".to_string());
                duplicates_rows.push(full_row);
            }
            processed_keys.insert(normalized_key.clone());
        }
    }

    // Process matches and singles
    for (normalized_key, left_indices) in &left_map {
        if processed_keys.contains(normalized_key) {
            continue;
        }
        let right_indices = right_map.get(normalized_key);

        if let Some(right_idxs) = right_indices {
            if right_idxs.len() == 1 && left_indices.len() == 1 {
                // Match
                let left_row = &input.left_rows[left_indices[0]];
                let right_row = &input.right_rows[right_idxs[0]];

                let mut result_row: Vec<String> = left_row.clone();
                result_row.extend(right_row.clone());

                // Find diff cols
                let mut diff_cols: Vec<String> = Vec::new();
                for (i, left_header) in input.left_headers.iter().enumerate() {
                    if let Some(right_idx) = input.right_headers.iter().position(|h| h == left_header) {
                        let left_val = left_row.get(i).map(|s| s.as_str()).unwrap_or("");
                        let right_val = right_row.get(right_idx).map(|s| s.as_str()).unwrap_or("");
                        if left_val != right_val {
                            diff_cols.push(left_header.clone());
                        }
                    }
                }

                result_row.push("both".to_string());
                result_row.push(diff_cols.join(","));
                result_row.push("0".to_string());
                result_rows.push(result_row);
            }
        } else {
            // Left only
            for &idx in left_indices {
                let mut row = input.left_rows[idx].clone();
                row.resize(result_headers.len() - 3, String::new());
                row.push("left_only".to_string());
                row.push(String::new());
                row.push("0".to_string());
                left_only_rows.push(row);
            }
        }
    }

    for (normalized_key, right_indices) in &right_map {
        if processed_keys.contains(normalized_key) {
            continue;
        }
        if !left_map.contains_key(normalized_key) {
            // Right only
            for &idx in right_indices {
                let mut row: Vec<String> = Vec::new();
                row.resize(input.left_headers.len(), String::new());
                row.extend(input.right_rows[idx].clone());
                row.resize(result_headers.len() - 3, String::new());
                row.push("right_only".to_string());
                row.push(String::new());
                row.push("0".to_string());
                right_only_rows.push(row);
            }
        }
    }

    let output = CompareOutput {
        result: TableData {
            headers: result_headers.clone(),
            rows: result_rows,
        },
        left_only: TableData {
            headers: result_headers.clone(),
            rows: left_only_rows,
        },
        right_only: TableData {
            headers: result_headers.clone(),
            rows: right_only_rows,
        },
        duplicates: TableData {
            headers: result_headers.clone(),
            rows: duplicates_rows,
        },
        log: vec![
            ("left_rows".to_string(), input.left_rows.len().to_string()),
            ("right_rows".to_string(), input.right_rows.len().to_string()),
            ("key_column".to_string(), input.key.clone()),
            ("trim".to_string(), input.options.trim.to_string()),
            ("case_insensitive".to_string(), input.options.case_insensitive.to_string()),
        ],
    };

    serde_json::to_string(&output).expect("Failed to serialize CompareOutput")
}

#[wasm_bindgen]
pub fn split_file(input_json: &str) -> String {
    let input: SplitInput = serde_json::from_str(input_json)
        .expect("Failed to parse SplitInput");
    
    let key_idx = input.headers.iter()
        .position(|h| h == &input.key)
        .expect("Key column not found in headers");

    let mut groups: std::collections::HashMap<String, Vec<Vec<String>>> = std::collections::HashMap::new();

    for row in input.rows {
        let key_value = row.get(key_idx)
            .map(|s| s.trim())
            .filter(|s| !s.is_empty())
            .unwrap_or("EMPTY")
            .to_string();
        groups.entry(key_value).or_insert_with(Vec::new).push(row);
    }

    let mut parts: Vec<SplitPart> = groups.into_iter()
        .map(|(key_value, rows)| {
            SplitPart {
                key_value: key_value.clone(),
                table: TableData {
                    headers: input.headers.clone(),
                    rows,
                },
            }
        })
        .collect();

    parts.sort_by(|a, b| a.key_value.cmp(&b.key_value));

    let output = SplitOutput { parts };
    serde_json::to_string(&output).expect("Failed to serialize SplitOutput")
}
