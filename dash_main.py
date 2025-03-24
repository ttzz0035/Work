import dash
import dash_bootstrap_components as dbc
from dash import html, dcc, Input, Output, State
import plotly.express as px

# Bootstrap & FontAwesome の読み込み
external_stylesheets = [
    dbc.themes.BOOTSTRAP,
    dbc.icons.FONT_AWESOME
]

app = dash.Dash(__name__, external_stylesheets=external_stylesheets)

# -----------------------------
# ダミーデータ
# -----------------------------
df_iris = px.data.iris()
df_gapminder = px.data.gapminder().query("country=='Japan' or country=='United States'")

# -----------------------------
# ページごとのレイアウト生成関数
# -----------------------------
def layout_home():
    return html.Div([
        html.H2("Home Page"),
        html.P("ここはトップページです。サイドバーから別のページに遷移できます。")
    ], style={"margin": "2rem"})

def layout_graph1():
    fig = px.scatter(df_iris, x="sepal_width", y="sepal_length",
                     color="species", title="Iris Scatter Plot")
    return html.Div([
        html.H2("Graph 1: Irisデータ（散布図）"),
        dcc.Graph(figure=fig)
    ], style={"margin": "2rem"})

def layout_graph2():
    fig = px.line(df_gapminder, x="year", y="lifeExp",
                  color="country", title="Life Expectancy over Time")
    return html.Div([
        html.H2("Graph 2: Gapminderデータ（折れ線グラフ）"),
        dcc.Graph(figure=fig)
    ], style={"margin": "2rem"})

def layout_settings2():
    return html.Div([
        html.H2("Settings Page2"),
        html.P("ここは設定ページです。自由に設定項目を追加できます。")
    ], style={"margin": "2rem"})

def layout_settings():
    return html.Div([
        html.H2("Settings Page"),
        html.P("ここは設定ページです。自由に設定項目を追加できます。")
    ], style={"margin": "2rem"})

# -----------------------------
# ページ情報のリスト
# -----------------------------
# ここで定義した情報が、サイドバー生成やページ表示のコールバックで共通利用されます。
page_list = [
    {"name": "Home", "href": "/", "layout": layout_home},
    {"name": "Graph 1", "href": "/graph1", "layout": layout_graph1},
    {"name": "Graph 2", "href": "/graph2", "layout": layout_graph2},
    {"name": "Settings2", "href": "/settings2", "layout": layout_settings2},
    {"name": "Settings", "href": "/settings", "layout": layout_settings},
]

# -----------------------------
# ページ情報リストから辞書マッピングを自動生成
# -----------------------------
page_map = {page["href"]: page["layout"] for page in page_list}

# -----------------------------
# サイドバー用メニュー生成関数
# -----------------------------
def generate_sidebar_menu(pages, exclude_hrefs=None):
    if exclude_hrefs is None:
        exclude_hrefs = []
    menu_items = []
    for page in pages:
        if page["href"] in exclude_hrefs:
            continue
        menu_items.append(
            dbc.ListGroupItem(
                dcc.Link(page["name"], href=page["href"]),
                className="border-0"
            )
        )
    return menu_items

# -----------------------------
# ヘッダー (Navbar)
# -----------------------------
header = dbc.Navbar(
    dbc.Container([
        dbc.NavbarBrand("My Dash App", href="/", className="me-auto"),
        dbc.NavItem(
            dbc.NavLink(
                html.I(className="fa fa-cog", style={"fontSize": "1.3em"}),
                href="/settings",
                id="settings-icon-link"
            )
        ),
    ], fluid=True),
    color="primary",
    dark=True,
    className="mb-2"
)

# -----------------------------
# サイドバー
# -----------------------------
sidebar = html.Div([
    # ハンバーガーアイコン部
    html.Div(
        dbc.Button(
            html.I(className="fa fa-bars"),
            id="sidebar-toggle",
            style={
                "background": "transparent",
                "border": "none",
                "fontSize": "1.4em",
                "margin": "0.5rem",
                "color": "#333",
            }
        ),
        style={"display": "flex", "alignItems": "center", "justifyContent": "flex-start"}
    ),
    # メニューリスト部（初期 style に display:block を設定）
    html.Div(
        id="menu-container",
        style={"display": "block"},
        children=dbc.ListGroup(
            generate_sidebar_menu(page_list, exclude_hrefs=["/settings"]),
            flush=True,
        )
    ),
],
    id="sidebar",
    className="bg-light",
    style={
        "position": "fixed",
        "top": "56px",         # ヘッダーの高さ
        "left": 0,
        "bottom": 0,
        "width": "250px",      # 展開時の幅
        "padding": "0.5rem",
        "transition": "width 0.3s, padding 0.3s",
        "overflow": "auto",
        "zIndex": 1000
    }
)

# -----------------------------
# メインコンテンツ
# -----------------------------
main_content = html.Div(
    id="page-content",
    style={
        "marginLeft": "250px",
        "transition": "margin-left 0.3s"
    }
)

# -----------------------------
# アプリ全体レイアウト
# -----------------------------
app.layout = html.Div([
    dcc.Location(id='url', refresh=False),
    header,
    sidebar,
    main_content
])

# -----------------------------
# (A) URLパスに応じたページ表示のコールバック
#    辞書マッピングを使うことで、リストと自動的に一致させています。
# -----------------------------
@app.callback(
    Output('page-content', 'children'),
    Input('url', 'pathname')
)
def display_page(pathname):
    layout_func = page_map.get(pathname, layout_home)
    return layout_func()

# -----------------------------
# (B) ハンバーガークリックでサイドバー折りたたみのコールバック
# -----------------------------
@app.callback(
    Output("sidebar", "style"),
    Output("menu-container", "style"),
    Output("page-content", "style"),
    Input("sidebar-toggle", "n_clicks"),
    State("sidebar", "style"),
    State("menu-container", "style"),
    State("page-content", "style"),
    prevent_initial_call=True
)
def toggle_sidebar(n_clicks, sidebar_style, menu_style, content_style):
    current_width = sidebar_style.get("width", "250px")
    if current_width == "250px":
        new_sidebar_style = sidebar_style.copy()
        new_sidebar_style["width"] = "50px"
        new_sidebar_style["padding"] = "0"
        new_menu_style = menu_style.copy() if menu_style else {"display": "none"}
        new_menu_style["display"] = "none"
        new_content_style = content_style.copy()
        new_content_style["marginLeft"] = "50px"
        return new_sidebar_style, new_menu_style, new_content_style
    else:
        new_sidebar_style = sidebar_style.copy()
        new_sidebar_style["width"] = "250px"
        new_sidebar_style["padding"] = "0.5rem"
        new_menu_style = menu_style.copy() if menu_style else {"display": "block"}
        new_menu_style["display"] = "block"
        new_content_style = content_style.copy()
        new_content_style["marginLeft"] = "250px"
        return new_sidebar_style, new_menu_style, new_content_style

if __name__ == "__main__":
    app.run(debug=True)
