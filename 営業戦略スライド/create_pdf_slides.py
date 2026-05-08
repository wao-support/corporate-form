import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
from matplotlib.backends.backend_pdf import PdfPages
import japanize_matplotlib
import numpy as np

# ============================================================
# カラー定義
# ============================================================
BG       = '#10101E'
ACCENT   = '#E83A3A'
ORANGE   = '#F5A623'
GREEN    = '#2ECC71'
WHITE    = '#FFFFFF'
LIGHT    = '#BBBBCC'
MUTED    = '#777799'
CARD     = '#1E1E35'
CARD2    = '#28284A'
DIVIDER  = '#444466'

W, H = 13.33, 7.5   # インチ（16:9）


def add_rect(ax, x, y, w, h, color, alpha=1.0, zorder=1,
             ec=None, lw=0, radius=0):
    if radius > 0:
        p = FancyBboxPatch((x, y), w, h,
                           boxstyle=f"round,pad=0,rounding_size={radius}",
                           fc=color, ec=ec or color, lw=lw,
                           alpha=alpha, zorder=zorder)
    else:
        p = patches.Rectangle((x, y), w, h,
                               fc=color, ec=ec or color, lw=lw,
                               alpha=alpha, zorder=zorder)
    ax.add_patch(p)

def add_text(ax, text, x, y, size=11, color=WHITE, bold=False,
             ha='left', va='top', zorder=5, alpha=1.0,
             wrap=False, width=None):
    weight = 'bold' if bold else 'normal'
    ax.text(x, y, text, fontsize=size, color=color, fontweight=weight,
            ha=ha, va=va, zorder=zorder, alpha=alpha,
            wrap=wrap)

def add_arrow(ax, x1, y1, x2, y2, color=WHITE, lw=1.5, zorder=4):
    ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle='->', color=color, lw=lw),
                zorder=zorder)


def make_fig():
    fig, ax = plt.subplots(figsize=(W, H))
    fig.patch.set_facecolor(BG)
    ax.set_facecolor(BG)
    ax.set_xlim(0, W)
    ax.set_ylim(0, H)
    ax.axis('off')
    ax.invert_yaxis()
    return fig, ax


# ============================================================
# ページ1：毎月リセット。この営業、いつまで続けられるか。
# ============================================================
def slide_crisis(ax):
    # 背景
    add_rect(ax, 0, 0, W, H, BG, zorder=0)

    # 上部アクセントバー
    add_rect(ax, 0, 0, W, 0.07, ACCENT, zorder=2)

    # 左縦ライン
    add_rect(ax, 0.3, 0.18, 0.06, H-0.3, ACCENT, zorder=2)

    # スライド番号バッジ
    add_rect(ax, 0.5, 0.18, 1.0, 0.30, ACCENT, zorder=3, radius=0.04)
    add_text(ax, '01 / 03', 1.0, 0.33, size=8, bold=True,
             color=WHITE, ha='center', va='center', zorder=4)

    # タイトル
    add_text(ax, '毎月リセット。この営業、いつまで続けられるか。',
             0.52, 0.60, size=19, bold=True, color=WHITE, zorder=4)

    # アンダーライン
    add_rect(ax, 0.52, 1.05, 12.5, 0.04, ACCENT, zorder=3)

    # ─── 左カラム ───────────────────────────────────────
    lines_left = [
        ('正直に言う。', WHITE, 11, False),
        ('', WHITE, 5, False),
        ('今のBP事業部の売上は、「たまたま」で成り立っている。', LIGHT, 10.5, False),
        ('大型案件が取れた月は達成。取れなかった月は未達。', LIGHT, 10.5, False),
        ('翌月の見込みは、月が変わるまでわからない。', LIGHT, 10.5, False),
        ('', WHITE, 5, False),
        ('これは「営業している」のではなく、', LIGHT, 10.5, False),
        ('「注文を待っている」だ。', ORANGE, 11, True),
    ]
    y = 1.25
    for text, color, size, bold in lines_left:
        if text:
            add_text(ax, text, 0.52, y, size=size, color=color, bold=bold, zorder=4)
        y += size * 0.018 + 0.04

    # 数字カード3枚
    cards = [
        ('月次目標', '3,000万円', '', CARD2, WHITE),
        ('現状見込み', '800〜900万円', '', CARD2, LIGHT),
        ('毎月の不足', '▲ 約2,000万円', 'どこからか来ることを祈っている', '#3A1010', ACCENT),
    ]
    card_w, card_h = 1.88, 1.05
    card_top = 3.05
    for i, (label, val, sub, bg, vc) in enumerate(cards):
        cx = 0.52 + i * (card_w + 0.14)
        ec = ACCENT if i == 2 else DIVIDER
        lw = 1.5 if i == 2 else 0.5
        add_rect(ax, cx, card_top, card_w, card_h, bg,
                 ec=ec, lw=lw, zorder=3, radius=0.06)
        add_text(ax, label, cx+0.12, card_top+0.18, size=9,
                 color=LIGHT, zorder=4)
        add_text(ax, val, cx+0.12, card_top+0.42, size=16,
                 bold=True, color=vc, zorder=4)
        if sub:
            add_text(ax, sub, cx+0.12, card_top+0.80, size=8.5,
                     color=MUTED, zorder=4)

    # ─── 右カラム ───────────────────────────────────────
    # 縦区切り線
    add_rect(ax, 6.6, 1.15, 0.03, 5.1, DIVIDER, zorder=2)

    add_text(ax, 'もっと深刻な事実がある。',
             6.78, 1.22, size=12, bold=True, color=ORANGE, zorder=4)

    intro_r = [
        '直近1年でアクティブな取引先は約 290社。',
        '過去10年、数百社の法人と関係を築いてきた。',
        '',
        'なのに——',
    ]
    y = 1.58
    for t in intro_r:
        add_text(ax, t, 6.78, y, size=10.5, color=LIGHT, zorder=4)
        y += 0.30

    # 問いかけブロック
    questions = [
        'その会社が年間いくら発注しているか、把握できているか？',
        '他の部署に発注担当者がいないか、確認したことがあるか？',
        'なぜ先月より注文が減ったのか、直接聞いたことがあるか？',
    ]
    for q in questions:
        add_rect(ax, 6.78, y, 6.25, 0.37, CARD2, zorder=3, radius=0.04)
        add_rect(ax, 6.78, y, 0.05, 0.37, ACCENT, zorder=4)
        add_text(ax, '▶  ' + q, 6.90, y+0.10, size=10,
                 color=LIGHT, zorder=4)
        y += 0.43

    add_text(ax, 'ほとんど、できていない。',
             6.78, y+0.10, size=14, bold=True, color=ACCENT, zorder=4)

    # ─── フッターバナー ─────────────────────────────────
    add_rect(ax, 0, 6.62, W, 0.88, ACCENT, zorder=3)
    add_text(ax,
             'このままでは売上は上がらない。チームも守れない。'
             '　営業チーム維持には月5,000万円規模の売上が必要だ。'
             '　今のやり方で、それが実現できるか——答えは全員わかっているはずだ。',
             W/2, 7.06, size=10.5, bold=True, color=WHITE,
             ha='center', va='center', zorder=4)


# ============================================================
# ページ2：「注文を待つ営業」を終わりにする。
# ============================================================
def slide_strategy(ax):
    add_rect(ax, 0, 0, W, H, BG, zorder=0)
    add_rect(ax, 0, 0, W, 0.07, ACCENT, zorder=2)
    add_rect(ax, 0.3, 0.18, 0.06, H-0.3, ACCENT, zorder=2)

    # バッジ
    add_rect(ax, 0.5, 0.18, 1.0, 0.30, ACCENT, zorder=3, radius=0.04)
    add_text(ax, '02 / 03', 1.0, 0.33, size=8, bold=True,
             color=WHITE, ha='center', va='center', zorder=4)

    # タイトル
    add_text(ax, '「注文を待つ営業」を終わりにする。',
             0.52, 0.60, size=21, bold=True, color=WHITE, zorder=4)
    add_rect(ax, 0.52, 1.06, 12.5, 0.04, ACCENT, zorder=3)

    # サブタイトル
    add_text(ax,
             '新しい戦略は一言で言える。既存顧客を、会社ごと握る。——これが「アカウント営業」だ。',
             0.52, 1.16, size=11.5, color=ORANGE, zorder=4)

    # ─── Before/After テーブル ──────────────────────────
    add_text(ax, '何が変わるのか', 0.52, 1.56, size=9.5,
             bold=True, color=MUTED, zorder=4)

    rows = [
        ('これまで', 'これから', True),
        ('担当者一人から注文をもらう', '会社全体の発注を継続的に取り続ける', False),
        ('単発スポットを積み上げる', 'ストック型の安定売上を構築する', False),
        ('注文が来るのを待つ', 'ニーズを掘り起こし、提案しにいく', False),
        ('案件単位で管理', '企業単位のアカウントで管理', False),
    ]
    rh = 0.37
    rt = 1.76
    for i, (before, after, is_hdr) in enumerate(rows):
        bg = ACCENT if is_hdr else (CARD if i % 2 == 1 else CARD2)
        fc = WHITE
        fs = 10 if not is_hdr else 10.5
        bw = 2.65

        add_rect(ax, 0.52, rt, bw, rh, bg, ec=DIVIDER, lw=0.3, zorder=3)
        add_text(ax, before, 0.65, rt+0.10, size=fs, bold=is_hdr, color=fc, zorder=4)

        if not is_hdr:
            add_text(ax, '→', 0.52+bw+0.07, rt+0.10, size=14,
                     bold=True, color=ACCENT, zorder=4)

        add_rect(ax, 0.52+bw+0.30, rt, bw, rh, bg, ec=DIVIDER, lw=0.3, zorder=3)
        add_text(ax, after, 0.52+bw+0.42, rt+0.10, size=fs,
                 bold=is_hdr, color=fc, zorder=4)
        rt += rh

    # ─── セグメント別戦術 ──────────────────────────────
    add_rect(ax, 6.6, 1.56, 0.03, 5.0, DIVIDER, zorder=2)

    add_text(ax, '顧客タイプ別のアプローチ', 6.75, 1.56,
             size=9.5, bold=True, color=MUTED, zorder=4)

    segs = [
        {
            'tag':    '集中型',
            'desc':   '特定の担当者1人に発注が集中している',
            'goal':   '目指すゴール：年間発注を握るパートナー化',
            'items':  ['年間見積を提示し、定例会を設定する',
                       '付加価値提供で競合を排除する',
                       'キーパーソンとの関係を組織単位に広げる'],
            'color':  ACCENT,
            'bg':     '#2A0808',
        },
        {
            'tag':    '分散型',
            'desc':   '社内の多くの社員がバラバラに発注している',
            'goal':   '目指すゴール：社内インフラ化（発注一本化）',
            'items':  ['利用状況をヒアリングし、課題を把握する',
                       '発注集約＋コスト削減の提案を持ち込む',
                       '利用ガイド作成など、社内展開を支援する'],
            'color':  ORANGE,
            'bg':     '#2A1E08',
        },
    ]
    seg_t = 1.80
    for seg in segs:
        add_rect(ax, 6.75, seg_t, 6.25, 1.75,
                 seg['bg'], ec=seg['color'], lw=1.2, zorder=3, radius=0.07)
        add_rect(ax, 6.75, seg_t, 1.1, 0.30,
                 seg['color'], zorder=4, radius=0.04)
        add_text(ax, seg['tag'], 7.30, seg_t+0.06, size=10,
                 bold=True, color=WHITE, ha='center', zorder=5)
        add_text(ax, seg['desc'], 7.95, seg_t+0.06, size=9.5,
                 color=LIGHT, zorder=4)
        add_text(ax, seg['goal'], 6.90, seg_t+0.40, size=10,
                 bold=True, color=seg['color'], zorder=4)
        for j, item in enumerate(seg['items']):
            add_text(ax, '・ ' + item, 6.90, seg_t+0.72+j*0.29,
                     size=9.5, color=LIGHT, zorder=4)
        seg_t += 1.88

    # ─── 実行メッセージ ──────────────────────────────
    add_rect(ax, 0.52, H-1.12, 12.5, 0.04, DIVIDER, zorder=2)
    add_text(ax,
             'やることは決まった。対象リストも出す。スクリプトも作る。あとは動くだけだ。',
             W/2, H-0.95, size=11, bold=False, color=ORANGE,
             ha='center', zorder=4)
    add_text(ax,
             '下期終了時に「あのとき変わってよかった」と言えるかどうかは、今週の行動で決まる。',
             W/2, H-0.62, size=11, bold=True, color=WHITE,
             ha='center', zorder=4)


# ============================================================
# ページ3：現状営業分析
# ============================================================
def slide_analysis(ax):
    add_rect(ax, 0, 0, W, H, BG, zorder=0)
    add_rect(ax, 0, 0, W, 0.07, ACCENT, zorder=2)
    add_rect(ax, 0.3, 0.18, 0.06, H-0.3, ACCENT, zorder=2)

    # バッジ
    add_rect(ax, 0.5, 0.18, 1.0, 0.30, ACCENT, zorder=3, radius=0.04)
    add_text(ax, '03 / 03', 1.0, 0.33, size=8, bold=True,
             color=WHITE, ha='center', va='center', zorder=4)

    # タグ
    add_text(ax, '現状分析', 1.65, 0.26, size=9, bold=True,
             color=ACCENT, zorder=4)

    # タイトル
    add_text(ax, '営業は「売上を作っている」のか？　データで検証した。',
             0.52, 0.60, size=18, bold=True, color=WHITE, zorder=4)
    add_rect(ax, 0.52, 1.06, 12.5, 0.04, ACCENT, zorder=3)

    # ─── 衝撃の一文 ────────────────────────────────────
    add_rect(ax, 1.4, 1.18, 10.5, 1.18,
             '#2A0808', ec=ACCENT, lw=1.8, zorder=3, radius=0.08)
    add_text(ax, '「営業が売上を作っているのか」を確かめに行ったら、',
             W/2, 1.44, size=11.5, color=LIGHT,
             ha='center', zorder=4)
    add_text(ax, '「営業は、売上を処理していた」',
             W/2, 1.78, size=22, bold=True, color=ACCENT,
             ha='center', zorder=4)
    add_rect(ax, 2.7, 1.75, 7.93, 0.04, '#662222', zorder=4)

    # ─── 3カラム Evidence ───────────────────────────────
    cols = [
        {
            'title': '① 初回活動の実態',
            'body':  ['多くの初回活動が「Re:メール」',
                      '（= 問い合わせへの返信）',
                      '',
                      '問い合わせが先にあり',
                      '営業はそれに対応している'],
            'conc':  '営業起点ではなく受動対応',
            'color': ACCENT,
        },
        {
            'title': '② 営業手法の分布',
            'body':  ['活動の大半がメール対応',
                      '電話（コール）は極めて少ない',
                      '',
                      'アウトバウンド営業は',
                      'ほぼ行われていない'],
            'conc':  '能動的な案件創出が弱い',
            'color': ORANGE,
        },
        {
            'title': '③ 売上の発生源',
            'body':  ['活動ログなし（#N/A）の',
                      '受注が一定数存在',
                      '',
                      'Web自然流入・紹介経由の',
                      '案件が相当数を占める'],
            'conc':  '営業直接創出の売上は不明',
            'color': GREEN,
        },
    ]
    cw, ch = 3.92, 2.72
    ct = 2.52
    for i, col in enumerate(cols):
        cx = 0.5 + i * (cw + 0.21)
        add_rect(ax, cx, ct, cw, ch,
                 CARD2, ec=DIVIDER, lw=0.5, zorder=3, radius=0.06)
        # ヘッダー
        add_rect(ax, cx, ct, cw, 0.34,
                 CARD, zorder=4, radius=0.06)
        add_text(ax, col['title'], cx+0.15, ct+0.09,
                 size=10, bold=True, color=col['color'], zorder=5)
        # 本文
        for j, line in enumerate(col['body']):
            add_text(ax, line, cx+0.18, ct+0.50+j*0.28,
                     size=10, color=LIGHT, zorder=4)
        # 結論バー
        add_rect(ax, cx, ct+ch-0.36, cw, 0.36,
                 '#1A1A30', zorder=4)
        add_rect(ax, cx, ct+ch-0.36, 0.05, 0.36,
                 col['color'], zorder=5)
        add_text(ax, col['conc'], cx+0.18, ct+ch-0.27,
                 size=9.5, bold=True, color=col['color'], zorder=5)

    # ─── 構造図 ─────────────────────────────────────────
    diag_t = 5.40
    add_rect(ax, 0.5, diag_t, 12.5, 0.98,
             '#18182C', ec=DIVIDER, lw=0.5, zorder=3)

    add_text(ax, '営業の役割：想定 vs 実態',
             0.68, diag_t+0.10, size=9, bold=True, color=MUTED, zorder=4)

    # フロー描画
    flow_data = [
        ('想定（あるべき姿）', GREEN,
         [('営業アプローチ', GREEN), ('案件創出', GREEN), ('受注', GREEN)], 0.65),
        ('実態（現実）', ACCENT,
         [('問い合わせ発生', MUTED), ('営業が対応', ORANGE), ('受注', ORANGE)], 7.0),
    ]
    bw2, bh2 = 1.62, 0.38
    for label, lc, boxes, sx in flow_data:
        add_text(ax, label, sx, diag_t+0.26, size=8.5,
                 bold=True, color=lc, zorder=4)
        for k, (bname, bc) in enumerate(boxes):
            bx = sx + k * (bw2 + 0.22)
            bg2 = '#224422' if bc == GREEN else ('#333328' if bc == ORANGE else '#283040')
            add_rect(ax, bx, diag_t+0.46, bw2, bh2,
                     bg2, ec=bc, lw=1.0, zorder=4, radius=0.04)
            add_text(ax, bname, bx+bw2/2, diag_t+0.52,
                     size=10, bold=True, color=bc,
                     ha='center', va='top', zorder=5)
            if k < 2:
                ax.annotate('', xy=(bx+bw2+0.22, diag_t+0.65),
                            xytext=(bx+bw2, diag_t+0.65),
                            arrowprops=dict(arrowstyle='->', color=bc, lw=1.2),
                            zorder=5)

    # VS バッジ
    add_rect(ax, 5.65, diag_t+0.43, 0.70, 0.42,
             ACCENT, zorder=5, radius=0.04)
    add_text(ax, 'VS', 6.0, diag_t+0.64, size=13,
             bold=True, color=WHITE, ha='center', va='center', zorder=6)

    # ─── フッター ────────────────────────────────────────
    add_rect(ax, 0, H-0.96, W, 0.96, '#1A0A0A', zorder=3)
    add_rect(ax, 0, H-0.96, W, 0.04, ACCENT, zorder=4)

    add_text(ax, '【示唆】', 0.52, H-0.80, size=9.5,
             bold=True, color=ACCENT, zorder=4)
    add_text(ax,
             '現状の営業組織は「案件創出装置」ではなく「案件処理装置」として機能している。'
             '　再現性ある売上拡大のためには、営業の役割を「対応」から「創出」へ再定義し、'
             '起点データの記録・KPIの再設計・アカウント営業への転換が不可欠だ。',
             1.52, H-0.80, size=10, color=LIGHT, zorder=4)


# ============================================================
# PDF出力
# ============================================================
out_path = '/Volumes/SSD-PGU3C/営業推進/営業戦略スライド/BP事業部_営業戦略.pdf'

with PdfPages(out_path) as pdf:
    for draw_func in [slide_crisis, slide_strategy, slide_analysis]:
        fig, ax = make_fig()
        draw_func(ax)
        plt.tight_layout(pad=0)
        pdf.savefig(fig, dpi=200, bbox_inches='tight', facecolor=BG)
        plt.close(fig)

print(f'Saved: {out_path}')
