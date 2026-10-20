#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <stdint.h>
#include <math.h>

#define MAX_STYLES 256
#define MAX_FILLS 128
#define MAX_FONTS 128
#define MAX_SHARED_STRINGS 10000
#define MAX_GRID_ROWS 1000
#define MAX_GRID_COLS 32

typedef struct {
    uint32_t clrBk;
    int is_valid;
} FillInfo;

typedef struct {
    uint32_t clrFg;
    int bold;
    int is_valid;
} FontInfo;

typedef struct {
    int fillId;
    int fontId;
    int numFmtId;
} XfInfo;

typedef struct {
    char* text;
} SharedStr;

typedef struct {
    char val[256];
    uint32_t clrBk;
    uint32_t clrFg;
    uint8_t flags; // bit 0: bold, bit 1: has_bk, bit 2: has_fg
} CellData;

typedef struct {
    uint16_t row;
    uint8_t col;
    uint8_t flags;
    uint32_t clrBk;
    uint32_t clrFg;
} __attribute__((packed)) StyleRecord;

static FillInfo g_fills[MAX_FILLS];
static int g_fill_count = 0;

static FontInfo g_fonts[MAX_FONTS];
static int g_font_count = 0;

static XfInfo g_xfs[MAX_STYLES];
static int g_xf_count = 0;

static char* g_shared_strings[MAX_SHARED_STRINGS];
static int g_shared_count = 0;

static CellData g_grid[MAX_GRID_ROWS][MAX_GRID_COLS];
static int g_max_row = 0;
static int g_max_col = 0;

// Sheet2 and Sheet3 lookup caches for cross-sheet references:
// Key: (sheet_idx * 10000 + row) * 32 + col -> double value
typedef struct {
    char text[64];
    double num;
    int has_num;
} SheetCell;

static SheetCell g_sheet2[500][16];
static SheetCell g_sheet3[500][16];

// Helper: Convert hex char to int
static int hex2int(char c) {
    if (c >= '0' && c <= '9') return c - '0';
    if (c >= 'a' && c <= 'f') return c - 'a' + 10;
    if (c >= 'A' && c <= 'F') return c - 'A' + 10;
    return 0;
}

// Convert 6 or 8 char hex string to Windows BGR COLORREF (0x00BBGGRR)
static uint32_t parse_hex_color(const char* hex, int len) {
    if (len == 8) {
        // AARRGGBB
        hex += 2;
        len = 6;
    }
    if (len != 6) return 0;
    int r = (hex2int(hex[0]) << 4) | hex2int(hex[1]);
    int g = (hex2int(hex[2]) << 4) | hex2int(hex[3]);
    int b = (hex2int(hex[4]) << 4) | hex2int(hex[5]);
    // Windows COLORREF is 0x00BBGGRR
    return (uint32_t)(r | (g << 8) | (b << 16));
}

// Read entire file into buffer
static char* read_file_to_memory(const char* path, size_t* out_len) {
    FILE* f = fopen(path, "rb");
    if (!f) return NULL;
    fseek(f, 0, SEEK_END);
    long len = ftell(f);
    fseek(f, 0, SEEK_SET);
    if (len <= 0) { fclose(f); return NULL; }
    char* buf = (char*)malloc(len + 1);
    if (!buf) { fclose(f); return NULL; }
    size_t rd = fread(buf, 1, len, f);
    buf[rd] = 0;
    fclose(f);
    if (out_len) *out_len = rd;
    return buf;
}

// Extract attribute value e.g. rgb="FF1B365D" -> out
static int extract_attr(const char* tag_start, const char* attr_name, char* out, int max_len) {
    char pattern[64];
    snprintf(pattern, sizeof(pattern), "%s=\"", attr_name);
    char* p = strstr(tag_start, pattern);
    if (!p) {
        snprintf(pattern, sizeof(pattern), "%s='", attr_name);
        p = strstr(tag_start, pattern);
        if (!p) return 0;
    }
    p += strlen(pattern);
    char* q = p;
    while (*q && *q != '"' && *q != '\'' && (q - p) < max_len - 1) {
        q++;
    }
    int len = (int)(q - p);
    memcpy(out, p, len);
    out[len] = 0;
    return 1;
}

// Convert UTF-8 string to ANSI CP_ACP (CP1251 on Russian systems), replacing special Unicode symbols
static void utf8_to_acp(const char* utf8, char* acp, int acp_max) {
    if (!utf8 || !acp || acp_max <= 0) return;
    wchar_t wbuf[2048];
    int wlen = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, wbuf, sizeof(wbuf)/sizeof(wchar_t));
    if (wlen > 0) {
        // Normalize special unicode characters to ASCII/CP1251 equivalents
        for (int i = 0; i < wlen; i++) {
            if (wbuf[i] == 0x2192) wbuf[i] = '>'; // ->
            else if (wbuf[i] == 0x2014) wbuf[i] = '-'; // em-dash
            else if (wbuf[i] == 0x2013) wbuf[i] = '-'; // en-dash
        }
        WideCharToMultiByte(CP_ACP, 0, wbuf, -1, acp, acp_max, "?", NULL);
    } else {
        strncpy(acp, utf8, acp_max - 1);
        acp[acp_max - 1] = 0;
    }
}

// Parse styles.xml
static void parse_styles_xml(const char* path) {
    size_t len = 0;
    char* buf = read_file_to_memory(path, &len);
    if (!buf) return;

    g_fill_count = 0;
    g_font_count = 0;
    g_xf_count = 0;

    // 1. Parse fills
    char* fills_tag = strstr(buf, "<fills");
    if (fills_tag) {
        char* p = fills_tag;
        while ((p = strstr(p, "<fill")) != NULL) {
            char* end_fill = strstr(p, "</fill>");
            char* next_fill = strstr(p + 5, "<fill");
            if (!end_fill && !next_fill) break;
            char* limit = end_fill ? end_fill : next_fill;

            // Search for fgColor rgb="..."
            char* fg = strstr(p, "<fgColor");
            if (fg && fg < limit) {
                char rgb[32] = {0};
                if (extract_attr(fg, "rgb", rgb, sizeof(rgb))) {
                    g_fills[g_fill_count].clrBk = parse_hex_color(rgb, strlen(rgb));
                    g_fills[g_fill_count].is_valid = 1;
                }
            }
            g_fill_count++;
            if (g_fill_count >= MAX_FILLS) break;
            p = limit + 6;
        }
    }

    // 2. Parse fonts
    char* fonts_tag = strstr(buf, "<fonts");
    if (fonts_tag) {
        char* p = fonts_tag;
        while ((p = strstr(p, "<font")) != NULL) {
            char* end_font = strstr(p, "</font>");
            char* next_font = strstr(p + 5, "<font");
            if (!end_font && !next_font) break;
            char* limit = end_font ? end_font : next_font;

            // Check bold <b/> or <b>
            if (strstr(p, "<b/>") && strstr(p, "<b/>") < limit) {
                g_fonts[g_font_count].bold = 1;
            } else if (strstr(p, "<b ") && strstr(p, "<b ") < limit) {
                g_fonts[g_font_count].bold = 1;
            }

            // Check color rgb="..."
            char* c_tag = strstr(p, "<color");
            if (c_tag && c_tag < limit) {
                char rgb[32] = {0};
                if (extract_attr(c_tag, "rgb", rgb, sizeof(rgb))) {
                    g_fonts[g_font_count].clrFg = parse_hex_color(rgb, strlen(rgb));
                    g_fonts[g_font_count].is_valid = 1;
                }
            }
            g_font_count++;
            if (g_font_count >= MAX_FONTS) break;
            p = limit + 6;
        }
    }

    // 3. Parse cellXfs
    char* xfs_tag = strstr(buf, "<cellXfs");
    if (xfs_tag) {
        char* p = xfs_tag;
        while ((p = strstr(p, "<xf ")) != NULL) {
            char fillId[16] = {0}, fontId[16] = {0}, numFmtId[16] = {0};
            extract_attr(p, "fillId", fillId, sizeof(fillId));
            extract_attr(p, "fontId", fontId, sizeof(fontId));
            extract_attr(p, "numFmtId", numFmtId, sizeof(numFmtId));

            g_xfs[g_xf_count].fillId = atoi(fillId);
            g_xfs[g_xf_count].fontId = atoi(fontId);
            g_xfs[g_xf_count].numFmtId = atoi(numFmtId);

            g_xf_count++;
            if (g_xf_count >= MAX_STYLES) break;
            p += 4;
        }
    }

    free(buf);
}

// Parse sharedStrings.xml
static void parse_shared_strings_xml(const char* path) {
    size_t len = 0;
    char* buf = read_file_to_memory(path, &len);
    if (!buf) return;

    g_shared_count = 0;
    char* p = buf;
    while ((p = strstr(p, "<si>")) != NULL) {
        char* end_si = strstr(p, "</si>");
        if (!end_si) break;

        char text[1024] = {0};
        int text_len = 0;

        // Collect all <t>...</t> inside this <si>
        char* t_ptr = p;
        while ((t_ptr = strstr(t_ptr, "<t")) != NULL && t_ptr < end_si) {
            char* t_content = strchr(t_ptr, '>');
            if (!t_content || t_content >= end_si) break;
            t_content++;
            char* t_end = strstr(t_content, "</t>");
            if (!t_end || t_end > end_si) break;

            int chunk = (int)(t_end - t_content);
            if (text_len + chunk < (int)sizeof(text) - 1) {
                memcpy(text + text_len, t_content, chunk);
                text_len += chunk;
                text[text_len] = 0;
            }
            t_ptr = t_end + 4;
        }

        char ansi_text[2048] = {0};
        utf8_to_acp(text, ansi_text, sizeof(ansi_text));
        g_shared_strings[g_shared_count++] = strdup(ansi_text);
        if (g_shared_count >= MAX_SHARED_STRINGS) break;
        p = end_si + 5;
    }

    free(buf);
}

// Convert column letters like "A" -> 1, "B" -> 2, "Z" -> 26, "AA" -> 27
static int col_letters_to_idx(const char* ref, int* out_row) {
    int col = 0;
    const char* p = ref;
    while (*p >= 'A' && *p <= 'Z') {
        col = col * 26 + (*p - 'A' + 1);
        p++;
    }
    if (out_row) {
        *out_row = atoi(p);
    }
    return col;
}

// Format number according to numFmtId
static void format_number(double val, int numFmtId, char* out, size_t out_sz) {
    if (numFmtId == 164 || numFmtId == 44) {
        // Currency: $#,##0.00
        if (fabs(val) >= 1000000.0) {
            double m = val / 1000000.0;
            long long whole = (long long)fabs(val);
            long long millions = whole / 1000000;
            long long thousands = (whole / 1000) % 1000;
            long long rem = whole % 1000;
            int cents = (int)(round(fabs(val) * 100.0)) % 100;
            snprintf(out, out_sz, "$%lld,%03lld,%03lld.%02d", millions, thousands, rem, cents);
        } else if (fabs(val) >= 1000.0) {
            long long whole = (long long)fabs(val);
            long long thousands = whole / 1000;
            long long rem = whole % 1000;
            int cents = (int)(round(fabs(val) * 100.0)) % 100;
            snprintf(out, out_sz, "$%lld,%03lld.%02d", thousands, rem, cents);
        } else {
            snprintf(out, out_sz, "$%.2f", val);
        }
    } else if (numFmtId == 165 || numFmtId == 9 || numFmtId == 10) {
        // Percentage: #,##0.0%
        double pct = val * 100.0;
        if (fabs(pct) >= 1000.0) {
            long long whole = (long long)fabs(pct);
            long long thousands = whole / 1000;
            long long rem = whole % 1000;
            int dec = (int)(round(fabs(pct) * 10.0)) % 10;
            snprintf(out, out_sz, "%lld,%03lld.%d%%", thousands, rem, dec);
        } else {
            snprintf(out, out_sz, "%.1f%%", pct);
        }
    } else if (numFmtId == 2 || numFmtId == 4) {
        // 0.00
        snprintf(out, out_sz, "%.2f", val);
    } else if (numFmtId == 1 || numFmtId == 3) {
        // Integer
        snprintf(out, out_sz, "%.0f", val);
    } else {
        // General or default
        if (val == floor(val) && fabs(val) < 1e9) {
            snprintf(out, out_sz, "%.0f", val);
        } else {
            snprintf(out, out_sz, "%.2f", val);
        }
    }
}

// Pre-scan secondary sheets (sheet2.xml, sheet3.xml) for fast cross-sheet references
static void prescan_sheet(const char* path, SheetCell cache[500][16]) {
    size_t len = 0;
    char* buf = read_file_to_memory(path, &len);
    if (!buf) return;

    char* p = buf;
    while ((p = strstr(p, "<c ")) != NULL) {
        char* end_c = strstr(p, "</c>");
        char* next_c = strstr(p + 3, "<c ");
        char* limit = end_c ? end_c + 4 : (next_c ? next_c : p + 200);

        char r_attr[32] = {0};
        if (!extract_attr(p, "r", r_attr, sizeof(r_attr))) {
            p += 3;
            continue;
        }

        int row = 0;
        int col = col_letters_to_idx(r_attr, &row);
        if (row < 500 && col < 16) {
            char* v_tag = strstr(p, "<v>");
            if (v_tag && v_tag < limit) {
                v_tag += 3;
                char* v_end = strstr(v_tag, "</v>");
                if (v_end && v_end < limit) {
                    int v_len = (int)(v_end - v_tag);
                    if (v_len > 0 && v_len < 60) {
                        memcpy(cache[row][col].text, v_tag, v_len);
                        cache[row][col].text[v_len] = 0;
                        cache[row][col].num = atof(cache[row][col].text);
                        cache[row][col].has_num = 1;
                    }
                }
            }
        }
        p = limit;
    }

    free(buf);
}

// Parse sheet1.xml into g_grid
static void parse_sheet_xml(const char* path) {
    size_t len = 0;
    char* buf = read_file_to_memory(path, &len);
    if (!buf) return;

    memset(g_grid, 0, sizeof(g_grid));
    g_max_row = 0;
    g_max_col = 0;

    char* p = buf;
    while ((p = strstr(p, "<c ")) != NULL) {
        char* end_c = strstr(p, "</c>");
        char* next_c = strstr(p + 3, "<c ");
        char* limit = end_c ? end_c + 4 : (next_c ? next_c : p + 400);

        char r_attr[32] = {0};
        char s_attr[16] = {0};
        char t_attr[16] = {0};

        extract_attr(p, "r", r_attr, sizeof(r_attr));
        extract_attr(p, "s", s_attr, sizeof(s_attr));
        extract_attr(p, "t", t_attr, sizeof(t_attr));

        int row_1 = 0;
        int col_1 = col_letters_to_idx(r_attr, &row_1);
        int r = row_1 - 1; // 0-based
        int c = col_1;     // 1-based (Col A = 1)

        if (r >= 0 && r < MAX_GRID_ROWS && c >= 0 && c < MAX_GRID_COLS) {
            if (r > g_max_row) g_max_row = r;
            if (c > g_max_col) g_max_col = c;

            // Apply style mapping
            int s_idx = atoi(s_attr);
            if (s_idx >= 0 && s_idx < g_xf_count) {
                int fillId = g_xfs[s_idx].fillId;
                int fontId = g_xfs[s_idx].fontId;

                if (fillId >= 0 && fillId < g_fill_count && g_fills[fillId].is_valid) {
                    uint32_t bk = g_fills[fillId].clrBk;
                    if (bk != 0) {
                        g_grid[r][c].clrBk = bk;
                        g_grid[r][c].flags |= 2; // has_bk
                    }
                }
                if (fontId >= 0 && fontId < g_font_count) {
                    if (g_fonts[fontId].bold) {
                        g_grid[r][c].flags |= 1; // bold
                    }
                    if (g_fonts[fontId].is_valid && g_fonts[fontId].clrFg != 0) {
                        g_grid[r][c].clrFg = g_fonts[fontId].clrFg;
                        g_grid[r][c].flags |= 4; // has_fg
                    }
                }
            }

            // Extract Value
            if (strcmp(t_attr, "s") == 0) {
                // Shared string index
                char* v_tag = strstr(p, "<v>");
                if (v_tag && v_tag < limit) {
                    int s_id = atoi(v_tag + 3);
                    if (s_id >= 0 && s_id < g_shared_count && g_shared_strings[s_id]) {
                        strncpy(g_grid[r][c].val, g_shared_strings[s_id], sizeof(g_grid[r][c].val) - 1);
                    }
                }
            } else if (strcmp(t_attr, "inlineStr") == 0) {
                // Inline string <is><t>...</t></is>
                char* t_tag = strstr(p, "<t");
                if (t_tag && t_tag < limit) {
                    char* t_val = strchr(t_tag, '>');
                    if (t_val && t_val < limit) {
                        t_val++;
                        char* t_end = strstr(t_val, "</t>");
                        if (t_end && t_end < limit) {
                            int v_len = (int)(t_end - t_val);
                            if (v_len > 0) {
                                char raw[512] = {0};
                                int copy_len = v_len < (int)sizeof(raw) - 1 ? v_len : (int)sizeof(raw) - 1;
                                memcpy(raw, t_val, copy_len);
                                raw[copy_len] = 0;
                                utf8_to_acp(raw, g_grid[r][c].val, sizeof(g_grid[r][c].val));
                            }
                        }
                    }
                }
            } else {
                // Formula or numeric value
                char* f_tag = strstr(p, "<f");
                char formula[128] = {0};
                if (f_tag && f_tag < limit) {
                    char* f_val = strchr(f_tag, '>');
                    if (f_val && f_val < limit) {
                        f_val++;
                        char* f_end = strstr(f_val, "</f>");
                        if (f_end && f_end < limit) {
                            int f_len = (int)(f_end - f_val);
                            if (f_len < (int)sizeof(formula) - 1) {
                                memcpy(formula, f_val, f_len);
                                formula[f_len] = 0;
                            }
                        }
                    }
                }

                char* v_tag = strstr(p, "<v>");
                char val_str[64] = {0};
                int has_val = 0;
                if (v_tag && v_tag < limit) {
                    v_tag += 3;
                    char* v_end = strstr(v_tag, "</v>");
                    if (v_end && v_end < limit) {
                        int v_len = (int)(v_end - v_tag);
                        if (v_len > 0 && v_len < (int)sizeof(val_str) - 1) {
                            memcpy(val_str, v_tag, v_len);
                            val_str[v_len] = 0;
                            has_val = 1;
                        }
                    }
                }

                int numFmtId = (s_idx >= 0 && s_idx < g_xf_count) ? g_xfs[s_idx].numFmtId : 0;

                if (has_val && val_str[0]) {
                    char* endptr = NULL;
                    double num = strtod(val_str, &endptr);
                    if (endptr && endptr != val_str && *endptr == 0) {
                        format_number(num, numFmtId, g_grid[r][c].val, sizeof(g_grid[r][c].val));
                    } else {
                        utf8_to_acp(val_str, g_grid[r][c].val, sizeof(g_grid[r][c].val));
                    }
                } else if (!has_val && formula[0]) {
                    // Check if references Sheet2 (Forex)
                    if (strstr(formula, "Форекс") || strstr(formula, "Forex")) {
                        char* ex = strchr(formula, '!');
                        if (ex) {
                            int srow = 0;
                            int scol = col_letters_to_idx(ex + 1, &srow);
                            if (srow < 500 && scol < 16 && g_sheet2[srow][scol].has_num) {
                                format_number(g_sheet2[srow][scol].num, numFmtId, g_grid[r][c].val, sizeof(g_grid[r][c].val));
                                has_val = 1;
                            }
                        }
                    } else if (strstr(formula, "Крипто") || strstr(formula, "24_7")) {
                        char* ex = strchr(formula, '!');
                        if (ex) {
                            int srow = 0;
                            int scol = col_letters_to_idx(ex + 1, &srow);
                            if (srow < 500 && scol < 16 && g_sheet3[srow][scol].has_num) {
                                format_number(g_sheet3[srow][scol].num, numFmtId, g_grid[r][c].val, sizeof(g_grid[r][c].val));
                                has_val = 1;
                            }
                        }
                    }
                    if (!has_val) {
                        snprintf(g_grid[r][c].val, sizeof(g_grid[r][c].val), "=%s", formula);
                    }
                }
            }
        }

        p = limit;
    }

    // Propagate background colors across merged cells (e.g. title banners A1:N1, section headers A4:E4)
    char* merge_tag = strstr(buf, "<mergeCells");
    if (merge_tag) {
        char* m = merge_tag;
        while ((m = strstr(m, "<mergeCell ")) != NULL) {
            char ref[64] = {0};
            if (extract_attr(m, "ref", ref, sizeof(ref))) {
                char* colon = strchr(ref, ':');
                if (colon) {
                    *colon = 0;
                    int r1 = 0, r2 = 0;
                    int c1 = col_letters_to_idx(ref, &r1);
                    int c2 = col_letters_to_idx(colon + 1, &r2);
                    r1--; r2--; // convert 1-based row to 0-based
                    if (r1 >= 0 && r1 < MAX_GRID_ROWS && c1 >= 1 && c1 < MAX_GRID_COLS) {
                        uint32_t bk = g_grid[r1][c1].clrBk;
                        uint8_t flags = g_grid[r1][c1].flags;
                        if (flags & 2) { // has_bk
                            for (int mr = r1; mr <= r2 && mr < MAX_GRID_ROWS; mr++) {
                                for (int mc = c1; mc <= c2 && mc < MAX_GRID_COLS; mc++) {
                                    g_grid[mr][mc].clrBk = bk;
                                    g_grid[mr][mc].flags |= 2;
                                }
                            }
                        }
                    }
                }
            }
            m += 10;
        }
    }

    free(buf);
}

// Convert extracted sheet to CSV and binary .style file
static int write_csv_and_style(const char* szCsvPath, const char* szStylePath) {
    FILE* fcsv = fopen(szCsvPath, "w");
    if (!fcsv) return -1;

    // Collect style records
    StyleRecord records[10000];
    int rec_count = 0;

    int max_c = g_max_col > 0 ? g_max_col : 10;
    if (max_c > 30) max_c = 30;

    for (int r = 0; r <= g_max_row; r++) {
        for (int c = 1; c <= max_c; c++) {
            CellData* cd = &g_grid[r][c];

            // Write CSV cell
            // If cell has comma, quotes or spaces, wrap in quotes
            int need_quotes = (strchr(cd->val, ',') != NULL || strchr(cd->val, '"') != NULL);
            if (need_quotes) {
                fputc('"', fcsv);
                for (char* p = cd->val; *p; p++) {
                    if (*p == '"') fputc('"', fcsv);
                    fputc(*p, fcsv);
                }
                fputc('"', fcsv);
            } else {
                fputs(cd->val, fcsv);
            }
            if (c < max_c) fputc(',', fcsv);

            // Record style
            if (cd->flags != 0 && rec_count < 10000) {
                records[rec_count].row = (uint16_t)r;
                records[rec_count].col = (uint8_t)c;
                records[rec_count].flags = cd->flags;
                records[rec_count].clrBk = cd->clrBk;
                records[rec_count].clrFg = cd->clrFg;
                rec_count++;
            }
        }
        fputs("\n", fcsv);
    }
    fclose(fcsv);

    // Write binary .style file
    FILE* fstyle = fopen(szStylePath, "wb");
    if (!fstyle) return -2;

    uint32_t magic = 0x53534C58; // 'XLSS'
    fwrite(&magic, 4, 1, fstyle);
    fwrite(&rec_count, 4, 1, fstyle);
    for (int i = 0; i < rec_count; i++) {
        fwrite(&records[i], sizeof(StyleRecord), 1, fstyle);
    }
    fclose(fstyle);

    return 0;
}

// Clean up temporary extracted directory
static void delete_temp_dir(const char* dir) {
    char cmd[512];
    snprintf(cmd, sizeof(cmd), "rmdir /s /q \"%s\"", dir);
    // Use Win32 CreateProcess to delete silently
    STARTUPINFOA si = { sizeof(si) };
    PROCESS_INFORMATION pi;
    char full_cmd[600];
    snprintf(full_cmd, sizeof(full_cmd), "cmd.exe /c %s", cmd);
    if (CreateProcessA(NULL, full_cmd, NULL, NULL, FALSE, CREATE_NO_WINDOW, NULL, NULL, &si, &pi)) {
        WaitForSingleObject(pi.hProcess, 3000);
        CloseHandle(pi.hProcess);
        CloseHandle(pi.hThread);
    }
}

// Main exported API:
// ConvertXlsxToCsvAndStyle
__declspec(dllexport) int ConvertXlsxToCsvAndStyle(const char* szXlsxPath, const char* szCsvPath, const char* szStylePath) {
    if (!szXlsxPath || !szCsvPath || !szStylePath) return -1;

    // 1. Create a unique temp folder
    wchar_t wTempPath[MAX_PATH];
    GetTempPathW(MAX_PATH, wTempPath);
    wchar_t wExtractDir[MAX_PATH];
    swprintf(wExtractDir, MAX_PATH, L"%lsxlsx_native_%lu", wTempPath, GetCurrentProcessId());
    CreateDirectoryW(wExtractDir, NULL);

    char szExtractDir[MAX_PATH];
    WideCharToMultiByte(CP_ACP, 0, wExtractDir, -1, szExtractDir, MAX_PATH, NULL, NULL);

    // 2. Unzip using tar.exe via CreateProcessW (handles Cyrillic paths and spaces seamlessly)
    wchar_t wXlsxPath[MAX_PATH];
    MultiByteToWideChar(CP_UTF8, 0, szXlsxPath, -1, wXlsxPath, MAX_PATH);
    if (GetFileAttributesW(wXlsxPath) == INVALID_FILE_ATTRIBUTES) {
        MultiByteToWideChar(CP_ACP, 0, szXlsxPath, -1, wXlsxPath, MAX_PATH);
    }

    wchar_t wCmd[MAX_PATH * 3];
    swprintf(wCmd, sizeof(wCmd)/sizeof(wchar_t), L"\"C:\\Windows\\System32\\tar.exe\" -xf \"%ls\" -C \"%ls\"", wXlsxPath, wExtractDir);

    STARTUPINFOW si = { sizeof(si) };
    PROCESS_INFORMATION pi;
    BOOL bSuccess = CreateProcessW(NULL, wCmd, NULL, NULL, FALSE, CREATE_NO_WINDOW, NULL, NULL, &si, &pi);
    if (!bSuccess) {
        // Fallback to relative tar.exe
        swprintf(wCmd, sizeof(wCmd)/sizeof(wchar_t), L"tar.exe -xf \"%ls\" -C \"%ls\"", wXlsxPath, wExtractDir);
        bSuccess = CreateProcessW(NULL, wCmd, NULL, NULL, FALSE, CREATE_NO_WINDOW, NULL, NULL, &si, &pi);
    }

    if (!bSuccess) {
        return -2; // tar.exe failed to spawn
    }

    WaitForSingleObject(pi.hProcess, 10000);
    DWORD exitCode = 0;
    GetExitCodeProcess(pi.hProcess, &exitCode);
    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);

    if (exitCode != 0) {
        delete_temp_dir(szExtractDir);
        return -3; // extraction failed
    }

    // 3. Paths inside extracted dir
    char pStyles[MAX_PATH], pShared[MAX_PATH], pSheet1[MAX_PATH], pSheet2[MAX_PATH], pSheet3[MAX_PATH];
    snprintf(pStyles, sizeof(pStyles), "%s\\xl\\styles.xml", szExtractDir);
    snprintf(pShared, sizeof(pShared), "%s\\xl\\sharedStrings.xml", szExtractDir);
    snprintf(pSheet1, sizeof(pSheet1), "%s\\xl\\worksheets\\sheet1.xml", szExtractDir);
    snprintf(pSheet2, sizeof(pSheet2), "%s\\xl\\worksheets\\sheet2.xml", szExtractDir);
    snprintf(pSheet3, sizeof(pSheet3), "%s\\xl\\worksheets\\sheet3.xml", szExtractDir);

    // Pre-scan secondary sheets for cross-sheet references
    memset(g_sheet2, 0, sizeof(g_sheet2));
    memset(g_sheet3, 0, sizeof(g_sheet3));
    prescan_sheet(pSheet2, g_sheet2);
    prescan_sheet(pSheet3, g_sheet3);

    // Parse styles & strings
    parse_styles_xml(pStyles);
    parse_shared_strings_xml(pShared);

    // Parse sheet1
    parse_sheet_xml(pSheet1);

    // Write CSV and .style
    int res = write_csv_and_style(szCsvPath, szStylePath);

    // Clean up strings
    for (int i = 0; i < g_shared_count; i++) {
        if (g_shared_strings[i]) free(g_shared_strings[i]);
    }
    g_shared_count = 0;

    // Delete temp directory
    delete_temp_dir(szExtractDir);

    return res;
}
