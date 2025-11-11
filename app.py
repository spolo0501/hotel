import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import numpy as np

# 頁面配置
st.set_page_config(
    page_title="W Hotel 客戶評價分析儀表板",
    page_icon="🏨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定義 CSS - 優化配色方案
st.markdown("""
<style>
    /* 主標題樣式 */
    .main-header {
        font-size: 2.8rem;
        font-weight: bold;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-align: center;
        padding: 1.5rem 0;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
    }

    /* 指標卡片優化 */
    .stMetric {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        padding: 1.2rem;
        border-radius: 0.8rem;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        border-left: 4px solid #667eea;
        transition: transform 0.3s;
    }

    .stMetric:hover {
        transform: translateY(-5px);
        box-shadow: 0 6px 12px rgba(0,0,0,0.15);
    }

    /* 區塊標題美化 */
    h2, h3 {
        color: #2d3748;
        border-bottom: 3px solid #667eea;
        padding-bottom: 0.5rem;
        margin-top: 2rem;
    }

    /* 側邊欄美化 */
    .css-1d391kg {
        background-color: #f8f9fa;
    }

    /* 標籤頁美化 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }

    .stTabs [data-baseweb="tab"] {
        background-color: #e2e8f0;
        border-radius: 0.5rem;
        padding: 0.5rem 1rem;
        font-weight: 500;
    }

    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
    }

    /* 資料表格美化 */
    .stDataFrame {
        border-radius: 0.5rem;
        overflow: hidden;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }

    /* 按鈕美化 */
    .stButton>button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 0.5rem;
        padding: 0.6rem 1.5rem;
        font-weight: 600;
        transition: all 0.3s;
    }

    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(102, 126, 234, 0.4);
    }

    /* 下載按鈕特殊樣式 */
    .stDownloadButton>button {
        background: linear-gradient(135deg, #48bb78 0%, #38a169 100%);
    }

    /* 選擇框美化 */
    .stSelectbox, .stMultiSelect {
        border-radius: 0.5rem;
    }

    /* 展開器美化 */
    .streamlit-expanderHeader {
        background-color: #edf2f7;
        border-radius: 0.5rem;
        font-weight: 600;
    }

    /* 警告/資訊框美化 */
    .stAlert {
        border-radius: 0.5rem;
        border-left: 4px solid #667eea;
    }
</style>
""", unsafe_allow_html=True)

# 載入數據
@st.cache_data
def load_data():
    df = pd.read_excel('chat_W_hotel.xlsx')
    df['date'] = pd.to_datetime(df['date'])
    df['year'] = df['date'].dt.year
    df['month'] = df['date'].dt.month
    df['year_month'] = df['date'].dt.to_period('M').astype(str)
    return df

# 主標題
st.markdown('<h1 class="main-header">🏨 W Hotel 客戶評價分析儀表板</h1>', unsafe_allow_html=True)

# 載入數據
try:
    df = load_data()

    # 側邊欄快速導航
    st.sidebar.header("🧭 快速導航")
    st.sidebar.markdown("""
    <style>
    .nav-link {
        display: block;
        padding: 0.5rem;
        margin: 0.2rem 0;
        background-color: #f0f2f6;
        border-radius: 0.3rem;
        text-decoration: none;
        color: #262730;
        transition: all 0.3s;
    }
    .nav-link:hover {
        background-color: #e0e2e6;
        transform: translateX(5px);
    }
    .nav-link-highlight {
        background-color: #ff4b4b;
        color: white;
        font-weight: bold;
    }
    .nav-link-highlight:hover {
        background-color: #ff3333;
    }
    </style>

    <a href="#kpi" class="nav-link">📊 關鍵指標</a>
    <a href="#trend" class="nav-link">📈 評價趨勢</a>
    <a href="#dimension-overview" class="nav-link">🎯 維度總覽</a>
    <a href="#dimension-compare" class="nav-link">🔀 維度比較 ✨</a>
    <a href="#wordcloud" class="nav-link">☁️ 關鍵詞雲 ✨</a>
    <a href="#distribution" class="nav-link">📊 評價分布</a>
    <a href="#drill-down" class="nav-link nav-link-highlight">🔍 維度深入分析 ⭐</a>
    <a href="#reviews" class="nav-link">💬 評論瀏覽</a>
    <a href="#download" class="nav-link">📥 資料下載</a>
    """, unsafe_allow_html=True)

    st.sidebar.markdown("---")

    # 側邊欄篩選器
    st.sidebar.header("📊 數據篩選")

    # 日期範圍篩選
    min_date = df['date'].min().date()
    max_date = df['date'].max().date()

    # 時間快捷選擇
    st.sidebar.markdown("**⏰ 時間快捷選擇**")
    time_preset = st.sidebar.radio(
        "選擇時間範圍",
        options=["自訂", "最近 30 天", "最近 3 個月", "最近 6 個月", "最近 1 年", "今年", "全部"],
        horizontal=False,
        label_visibility="collapsed"
    )

    from datetime import timedelta
    today = datetime.now().date()

    if time_preset == "最近 30 天":
        start_date = max(today - timedelta(days=30), min_date)
        end_date = max_date
    elif time_preset == "最近 3 個月":
        start_date = max(today - timedelta(days=90), min_date)
        end_date = max_date
    elif time_preset == "最近 6 個月":
        start_date = max(today - timedelta(days=180), min_date)
        end_date = max_date
    elif time_preset == "最近 1 年":
        start_date = max(today - timedelta(days=365), min_date)
        end_date = max_date
    elif time_preset == "今年":
        start_date = max(datetime(today.year, 1, 1).date(), min_date)
        end_date = max_date
    elif time_preset == "全部":
        start_date = min_date
        end_date = max_date
    else:  # 自訂
        date_range = st.sidebar.date_input(
            "自訂日期範圍",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date
        )
        if len(date_range) == 2:
            start_date, end_date = date_range
        else:
            start_date, end_date = min_date, max_date

    # 星級篩選
    star_options = sorted(df['star'].dropna().unique())
    selected_stars = st.sidebar.multiselect(
        "選擇星級",
        options=star_options,
        default=star_options
    )

    # 情感篩選
    sentiment_map = {-1.0: '負面', 0.0: '中性', 1.0: '正面'}
    selected_sentiments = st.sidebar.multiselect(
        "選擇情感",
        options=list(sentiment_map.values()),
        default=list(sentiment_map.values())
    )

    # 反向映射情感值
    sentiment_reverse_map = {'負面': -1.0, '中性': 0.0, '正面': 1.0}
    selected_sentiment_values = [sentiment_reverse_map[s] for s in selected_sentiments]

    # 應用篩選
    filtered_df = df[
        (df['date'].dt.date >= start_date) &
        (df['date'].dt.date <= end_date) &
        (df['star'].isin(selected_stars)) &
        (df['sentiment'].isin(selected_sentiment_values))
    ]

    st.sidebar.markdown(f"**篩選後數據量**: {len(filtered_df)} / {len(df)} 筆")

    # KPI 指標區
    st.markdown('<a id="kpi"></a>', unsafe_allow_html=True)
    st.markdown("---")
    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        st.metric(
            label="📝 總評論數",
            value=f"{len(filtered_df):,}"
        )

    with col2:
        avg_star = filtered_df['star'].mean()
        st.metric(
            label="⭐ 平均星級",
            value=f"{avg_star:.2f}"
        )

    with col3:
        positive_pct = (filtered_df['sentiment'] == 1.0).sum() / len(filtered_df) * 100 if len(filtered_df) > 0 else 0
        st.metric(
            label="😊 正面評價比例",
            value=f"{positive_pct:.1f}%"
        )

    with col4:
        negative_pct = (filtered_df['sentiment'] == -1.0).sum() / len(filtered_df) * 100 if len(filtered_df) > 0 else 0
        st.metric(
            label="😞 負面評價比例",
            value=f"{negative_pct:.1f}%"
        )

    with col5:
        date_span = (filtered_df['date'].max() - filtered_df['date'].min()).days
        st.metric(
            label="📅 時間跨度",
            value=f"{date_span} 天"
        )

    st.markdown("---")

    # 第一行：時間趨勢圖
    st.markdown('<a id="trend"></a>', unsafe_allow_html=True)
    st.subheader("📈 評價趨勢分析")

    tab1, tab2, tab3 = st.tabs(["月度趨勢", "年度趨勢", "情感趨勢"])

    with tab1:
        # 月度趨勢
        monthly_data = filtered_df.groupby('year_month').agg({
            'star': 'mean',
            'sentiment': 'mean',
            'text': 'count'
        }).reset_index()
        monthly_data.columns = ['年月', '平均星級', '平均情感分數', '評論數']

        fig1 = go.Figure()
        fig1.add_trace(go.Scatter(
            x=monthly_data['年月'],
            y=monthly_data['平均星級'],
            mode='lines+markers',
            name='平均星級',
            line=dict(color='#667eea', width=3),
            marker=dict(size=8, color='#764ba2')
        ))

        fig1.add_trace(go.Bar(
            x=monthly_data['年月'],
            y=monthly_data['評論數'],
            name='評論數',
            yaxis='y2',
            opacity=0.3,
            marker_color='lightgray'
        ))

        fig1.update_layout(
            title='月度平均星級趨勢',
            xaxis_title='年月',
            yaxis_title='平均星級',
            yaxis2=dict(
                title='評論數',
                overlaying='y',
                side='right'
            ),
            hovermode='x unified',
            height=400,
            showlegend=True
        )

        st.plotly_chart(fig1, use_container_width=True)

    with tab2:
        # 年度趨勢
        yearly_data = filtered_df.groupby('year').agg({
            'star': 'mean',
            'text': 'count',
            'sentiment': 'mean'
        }).reset_index()
        yearly_data.columns = ['年份', '平均星級', '評論數', '平均情感分數']

        fig2 = go.Figure()
        fig2.add_trace(go.Bar(
            x=yearly_data['年份'],
            y=yearly_data['平均星級'],
            name='平均星級',
            text=yearly_data['平均星級'].round(2),
            textposition='auto',
            marker=dict(
                color=yearly_data['平均星級'],
                colorscale='Viridis',
                showscale=False
            )
        ))

        fig2.update_layout(
            title='年度平均星級',
            xaxis_title='年份',
            yaxis_title='平均星級',
            height=400
        )

        st.plotly_chart(fig2, use_container_width=True)

        # 顯示年度統計表
        st.dataframe(yearly_data, use_container_width=True)

    with tab3:
        # 情感分布趨勢（改為百分比堆疊圖）
        sentiment_time = filtered_df.groupby(['year_month', 'sentiment']).size().reset_index(name='count')

        # 計算每個月的總數和百分比
        total_by_month = sentiment_time.groupby('year_month')['count'].sum().reset_index()
        total_by_month.columns = ['year_month', 'total']
        sentiment_time = sentiment_time.merge(total_by_month, on='year_month')
        sentiment_time['percentage'] = (sentiment_time['count'] / sentiment_time['total'] * 100).round(1)
        sentiment_time['sentiment_label'] = sentiment_time['sentiment'].map(sentiment_map)

        fig3 = px.area(
            sentiment_time,
            x='year_month',
            y='percentage',
            color='sentiment_label',
            title='情感分布時間趨勢（百分比）',
            labels={'year_month': '年月', 'percentage': '百分比 (%)', 'sentiment_label': '情感'},
            color_discrete_map={'正面': '#48bb78', '中性': '#ed8936', '負面': '#f56565'},
            groupnorm='percent'  # 堆疊百分比模式
        )

        fig3.update_layout(
            height=400,
            yaxis=dict(range=[0, 100], ticksuffix='%'),
            hovermode='x unified'
        )

        st.plotly_chart(fig3, use_container_width=True)

        # 顯示統計摘要
        col1, col2, col3 = st.columns(3)
        with col1:
            avg_positive = sentiment_time[sentiment_time['sentiment_label'] == '正面']['percentage'].mean()
            st.metric("平均正面比例", f"{avg_positive:.1f}%")
        with col2:
            avg_neutral = sentiment_time[sentiment_time['sentiment_label'] == '中性']['percentage'].mean()
            st.metric("平均中性比例", f"{avg_neutral:.1f}%")
        with col3:
            avg_negative = sentiment_time[sentiment_time['sentiment_label'] == '負面']['percentage'].mean()
            st.metric("平均負面比例", f"{avg_negative:.1f}%")

    st.markdown("---")

    # 第二行：評分維度分析
    st.markdown('<a id="dimension-overview"></a>', unsafe_allow_html=True)
    st.subheader("🎯 各維度評分分析")

    col1, col2 = st.columns(2)

    with col1:
        # 各維度平均分數
        dimensions = [
            'r_sentiment.Staff Service',
            'r_sentiment.Location',
            'r_sentiment.Room & Bathroom Quality',
            'r_sentiment.Environment',
            'r_sentiment.Facilities',
            'r_sentiment.Food & Beverage',
            'r_sentiment.Value'
        ]

        dimension_names = [
            '員工服務',
            '地點位置',
            '房間浴室品質',
            '環境',
            '設施',
            '餐飲',
            '性價比'
        ]

        avg_scores = []
        for dim in dimensions:
            avg_scores.append(filtered_df[dim].mean())

        dimension_df = pd.DataFrame({
            '維度': dimension_names,
            '平均分數': avg_scores
        }).sort_values('平均分數', ascending=True)

        fig4 = px.bar(
            dimension_df,
            x='平均分數',
            y='維度',
            orientation='h',
            title='各維度平均情感分數',
            color='平均分數',
            color_continuous_scale='RdYlGn',
            text='平均分數'
        )

        fig4.update_traces(texttemplate='%{text:.2f}', textposition='outside')
        fig4.update_layout(height=400, showlegend=False)
        st.plotly_chart(fig4, use_container_width=True)

    with col2:
        # 雷達圖
        radar_df = dimension_df[dimension_df['平均分數'].notna()]

        fig5 = go.Figure()

        fig5.add_trace(go.Scatterpolar(
            r=radar_df['平均分數'].tolist() + [radar_df['平均分數'].tolist()[0]],
            theta=radar_df['維度'].tolist() + [radar_df['維度'].tolist()[0]],
            fill='toself',
            name='平均分數',
            line=dict(color='#667eea', width=2),
            fillcolor='rgba(102, 126, 234, 0.4)'
        ))

        fig5.update_layout(
            polar=dict(
                radialaxis=dict(
                    visible=True,
                    range=[-1, 1]
                )
            ),
            showlegend=False,
            title='各維度評分雷達圖',
            height=400
        )

        st.plotly_chart(fig5, use_container_width=True)

    st.markdown("---")

    # 維度比較功能
    st.markdown('<a id="dimension-compare"></a>', unsafe_allow_html=True)
    st.subheader("🔀 維度比較分析")
    st.markdown("*選擇多個維度進行橫向比較*")

    # 重新定義維度映射（確保作用域正確）
    dimension_map_compare = {
        '員工服務': 'r_sentiment.Staff Service',
        '地點位置': 'r_sentiment.Location',
        '房間浴室品質': 'r_sentiment.Room & Bathroom Quality',
        '環境': 'r_sentiment.Environment',
        '設施': 'r_sentiment.Facilities',
        '餐飲': 'r_sentiment.Food & Beverage',
        '性價比': 'r_sentiment.Value'
    }

    # 維度選擇器（多選）
    compare_dimensions = st.multiselect(
        "選擇要比較的維度（建議 2-4 個）",
        options=list(dimension_map_compare.keys()),
        default=list(dimension_map_compare.keys())[:3]
    )

    if len(compare_dimensions) >= 2:
        # 準備比較數據
        compare_data = []
        for dim_name in compare_dimensions:
            dim_col = dimension_map_compare[dim_name]
            avg_score = filtered_df[dim_col].mean()
            positive_rate = (filtered_df[dim_col] > 0).sum() / filtered_df[dim_col].notna().sum() * 100 if filtered_df[dim_col].notna().sum() > 0 else 0
            count = filtered_df[dim_col].notna().sum()
            compare_data.append({
                '維度': dim_name,
                '平均分數': avg_score,
                '正面評論比例': positive_rate,
                '評論數': count
            })

        compare_df = pd.DataFrame(compare_data)

        col1, col2 = st.columns(2)

        with col1:
            # 平均分數比較
            fig_compare1 = go.Figure(data=[
                go.Bar(
                    x=compare_df['維度'],
                    y=compare_df['平均分數'],
                    text=compare_df['平均分數'].round(2),
                    textposition='auto',
                    marker=dict(
                        color=compare_df['平均分數'],
                        colorscale='RdYlGn',
                        showscale=True,
                        colorbar=dict(title="分數")
                    )
                )
            ])

            fig_compare1.update_layout(
                title='各維度平均分數比較',
                xaxis_title='維度',
                yaxis_title='平均分數',
                yaxis=dict(range=[-1, 1]),
                height=400
            )

            st.plotly_chart(fig_compare1, use_container_width=True)

        with col2:
            # 正面評論比例比較
            fig_compare2 = go.Figure(data=[
                go.Bar(
                    x=compare_df['維度'],
                    y=compare_df['正面評論比例'],
                    text=compare_df['正面評論比例'].round(1).astype(str) + '%',
                    textposition='auto',
                    marker=dict(
                        color=compare_df['正面評論比例'],
                        colorscale='Greens',
                        showscale=True,
                        colorbar=dict(title="比例 (%)")
                    )
                )
            ])

            fig_compare2.update_layout(
                title='各維度正面評論比例',
                xaxis_title='維度',
                yaxis_title='正面評論比例 (%)',
                yaxis=dict(range=[0, 100]),
                height=400
            )

            st.plotly_chart(fig_compare2, use_container_width=True)

        # 顯示詳細比較表格
        st.markdown("**詳細比較數據**")
        compare_df['平均分數'] = compare_df['平均分數'].round(3)
        compare_df['正面評論比例'] = compare_df['正面評論比例'].round(1).astype(str) + '%'
        st.dataframe(compare_df, use_container_width=True, hide_index=True)

    else:
        st.info("💡 請至少選擇 2 個維度進行比較")

    st.markdown("---")

    # 詞雲圖區域
    st.markdown('<a id="wordcloud"></a>', unsafe_allow_html=True)
    st.subheader("☁️ 評論關鍵詞雲")
    st.markdown("*查看評論中最常出現的詞彙*")

    wordcloud_sentiment = st.radio(
        "選擇要分析的情感類型",
        options=['全部', '正面', '中性', '負面'],
        horizontal=True
    )

    # 根據選擇篩選文字
    if wordcloud_sentiment == '正面':
        wordcloud_df = filtered_df[filtered_df['sentiment'] == 1.0]
    elif wordcloud_sentiment == '中性':
        wordcloud_df = filtered_df[filtered_df['sentiment'] == 0.0]
    elif wordcloud_sentiment == '負面':
        wordcloud_df = filtered_df[filtered_df['sentiment'] == -1.0]
    else:
        wordcloud_df = filtered_df

    # 合併所有評論文字
    all_text = ' '.join(wordcloud_df['text'].dropna().astype(str))

    if len(all_text) > 0:
        # 簡易詞頻統計（中文分詞需要 jieba，這裡先用簡單的字詞統計）
        import re
        from collections import Counter

        # 移除標點符號和數字
        text_cleaned = re.sub(r'[^\w\s]', ' ', all_text)
        text_cleaned = re.sub(r'\d+', '', text_cleaned)

        # 簡單的詞頻統計（這裡統計 2-4 個字的詞）
        words = []
        for length in [2, 3, 4]:
            for i in range(len(text_cleaned) - length + 1):
                word = text_cleaned[i:i+length]
                if word.strip() and not word.isspace():
                    words.append(word.strip())

        # 統計詞頻
        word_freq = Counter(words)
        # 過濾停用詞（常見但無意義的詞）
        stop_words = {'的', '了', '和', '是', '在', '有', '我', '就', '不', '也', '都', '這', '那', '要', '會', '可', '能', '但', '很', '還', '沒', '說', '而', '到', '去', '對', '與', '及', '以', '被', '給', '把', '讓', '為', '從', '向', '於', '比', '讓我', '我們', '你們', '他們', '這個', '那個', '什麼', '如果', '因為', '所以', '雖然', '然而', '當然', '可以', '應該', '可能', '一定'}
        word_freq = {k: v for k, v in word_freq.items() if k not in stop_words and v >= 3}

        # 取前 30 個高頻詞
        top_words = dict(sorted(word_freq.items(), key=lambda x: x[1], reverse=True)[:30])

        if top_words:
            # 使用柱狀圖顯示詞頻（替代詞雲）
            words_df = pd.DataFrame(list(top_words.items()), columns=['詞彙', '出現次數'])
            words_df = words_df.sort_values('出現次數', ascending=True).tail(20)

            fig_words = go.Figure(data=[
                go.Bar(
                    y=words_df['詞彙'],
                    x=words_df['出現次數'],
                    orientation='h',
                    text=words_df['出現次數'],
                    textposition='auto',
                    marker=dict(
                        color=words_df['出現次數'],
                        colorscale='Viridis',
                        showscale=False
                    )
                )
            ])

            fig_words.update_layout(
                title=f'前 20 名高頻詞彙 - {wordcloud_sentiment}評論',
                xaxis_title='出現次數',
                yaxis_title='詞彙',
                height=600,
                showlegend=False
            )

            st.plotly_chart(fig_words, use_container_width=True)

            # 顯示完整詞頻表
            with st.expander("📋 查看完整詞頻列表"):
                full_words_df = pd.DataFrame(list(top_words.items()), columns=['詞彙', '出現次數'])
                full_words_df = full_words_df.sort_values('出現次數', ascending=False)
                st.dataframe(full_words_df, use_container_width=True, hide_index=True)

        else:
            st.info("📝 沒有足夠的詞彙數據生成詞頻統計（詞彙至少需出現 3 次）")
    else:
        st.warning("⚠️ 沒有符合條件的評論文字")

    st.markdown("---")

    # 第三行：星級與情感分布
    st.markdown('<a id="distribution"></a>', unsafe_allow_html=True)
    st.subheader("📊 評價分布分析")

    col1, col2 = st.columns(2)

    with col1:
        # 星級分布
        star_dist = filtered_df['star'].value_counts().sort_index()

        fig6 = go.Figure()
        fig6.add_trace(go.Bar(
            x=star_dist.index,
            y=star_dist.values,
            text=star_dist.values,
            textposition='auto',
            marker=dict(
                color=['#f56565', '#fc8181', '#fbd38d', '#4299e1', '#48bb78'],
                line=dict(color='white', width=2)
            )
        ))

        fig6.update_layout(
            title='星級分布',
            xaxis_title='星級',
            yaxis_title='評論數',
            height=400
        )

        st.plotly_chart(fig6, use_container_width=True)

    with col2:
        # 情感分布圓餅圖
        sentiment_dist = filtered_df['sentiment'].value_counts()
        sentiment_labels = [sentiment_map.get(k, '未知') for k in sentiment_dist.index]

        fig7 = go.Figure(data=[go.Pie(
            labels=sentiment_labels,
            values=sentiment_dist.values,
            hole=0.4,
            marker=dict(
                colors=['#f56565', '#ed8936', '#48bb78'],
                line=dict(color='white', width=2)
            ),
            textfont=dict(size=14, color='white', family='Arial')
        )])

        fig7.update_layout(
            title='情感分布',
            height=400
        )

        st.plotly_chart(fig7, use_container_width=True)

    st.markdown("---")

    # 維度鑽取分析區（類似 PowerBI 功能）
    st.markdown('<a id="drill-down"></a>', unsafe_allow_html=True)
    st.subheader("🔍 維度深入分析（Drill-down）⭐")
    st.markdown("*點選維度查看該面向的詳細評論與情感分布*")

    # 維度映射
    dimension_mapping = {
        '員工服務': {
            'sentiment_col': 'r_sentiment.Staff Service',
            'reasons_col': 'reasons.Staff Service'
        },
        '地點位置': {
            'sentiment_col': 'r_sentiment.Location',
            'reasons_col': 'reasons.Location'
        },
        '房間浴室品質': {
            'sentiment_col': 'r_sentiment.Room & Bathroom Quality',
            'reasons_col': 'reasons.Room & Bathroom Quality'
        },
        '環境': {
            'sentiment_col': 'r_sentiment.Environment',
            'reasons_col': 'reasons.Environment'
        },
        '設施': {
            'sentiment_col': 'r_sentiment.Facilities',
            'reasons_col': 'reasons.Facilities'
        },
        '餐飲': {
            'sentiment_col': 'r_sentiment.Food & Beverage',
            'reasons_col': 'reasons.Food & Beverage'
        },
        '性價比': {
            'sentiment_col': 'r_sentiment.Value',
            'reasons_col': 'reasons.Value'
        }
    }

    # 選擇要分析的維度
    selected_dimension = st.selectbox(
        "🎯 選擇要深入分析的維度",
        options=list(dimension_mapping.keys()),
        index=0
    )

    # 獲取選定維度的欄位
    sentiment_col = dimension_mapping[selected_dimension]['sentiment_col']
    reasons_col = dimension_mapping[selected_dimension]['reasons_col']

    # 篩選出該維度有資料的評論
    dimension_df = filtered_df[filtered_df[sentiment_col].notna()].copy()

    if len(dimension_df) > 0:
        col1, col2, col3 = st.columns(3)

        with col1:
            st.metric(
                label=f"📊 {selected_dimension} - 總評論數",
                value=f"{len(dimension_df):,}"
            )

        with col2:
            avg_sentiment = dimension_df[sentiment_col].mean()
            sentiment_emoji = "😊" if avg_sentiment > 0.3 else "😞" if avg_sentiment < -0.3 else "😐"
            st.metric(
                label=f"{sentiment_emoji} 平均情感分數",
                value=f"{avg_sentiment:.2f}"
            )

        with col3:
            positive_count = (dimension_df[sentiment_col] > 0).sum()
            positive_rate = positive_count / len(dimension_df) * 100
            st.metric(
                label="✅ 正面評論比例",
                value=f"{positive_rate:.1f}%"
            )

        st.markdown("---")

        # 情感分布與評論內容
        col1, col2 = st.columns([1, 2])

        with col1:
            # 該維度的情感分布圓餅圖
            sentiment_counts = dimension_df[sentiment_col].apply(
                lambda x: '正面' if x > 0 else '負面' if x < 0 else '中性'
            ).value_counts()

            fig_dim = go.Figure(data=[go.Pie(
                labels=sentiment_counts.index,
                values=sentiment_counts.values,
                hole=0.4,
                marker=dict(
                    colors=['#48bb78', '#ed8936', '#f56565'],
                    line=dict(color='white', width=2)
                ),
                textfont=dict(size=13, color='white', family='Arial')
            )])

            fig_dim.update_layout(
                title=f'{selected_dimension} 情感分布',
                height=350
            )

            st.plotly_chart(fig_dim, use_container_width=True)

            # 情感篩選器（針對該維度）
            dim_sentiment_filter = st.radio(
                "篩選情感",
                options=['全部', '正面', '中性', '負面'],
                horizontal=True
            )

        with col2:
            # 根據情感篩選
            if dim_sentiment_filter == '正面':
                filtered_dim_df = dimension_df[dimension_df[sentiment_col] > 0]
            elif dim_sentiment_filter == '中性':
                filtered_dim_df = dimension_df[dimension_df[sentiment_col] == 0]
            elif dim_sentiment_filter == '負面':
                filtered_dim_df = dimension_df[dimension_df[sentiment_col] < 0]
            else:
                filtered_dim_df = dimension_df

            st.markdown(f"**顯示 {len(filtered_dim_df)} 筆評論**")

            # 顯示該維度的評論摘要和完整評論
            display_dim_df = filtered_dim_df[[
                'date', 'name', 'star', sentiment_col, reasons_col, 'text'
            ]].sort_values('date', ascending=False).head(10).copy()

            # 格式化顯示
            display_dim_df.columns = ['日期', '姓名', '星級', '情感分數', f'{selected_dimension}相關評論', '完整評論']
            display_dim_df['情感分數'] = display_dim_df['情感分數'].round(2)

            st.dataframe(
                display_dim_df,
                use_container_width=True,
                height=350,
                column_config={
                    '日期': st.column_config.DateColumn('日期', format='YYYY-MM-DD'),
                    '情感分數': st.column_config.NumberColumn(
                        '情感分數',
                        format='%.2f',
                        help='-1 (負面) ~ 1 (正面)'
                    )
                }
            )

        # 詳細評論展開區
        with st.expander(f"💬 查看 {selected_dimension} 的詳細評論內容"):
            for idx, row in filtered_dim_df.head(5).iterrows():
                sentiment_color = "🟢" if row[sentiment_col] > 0 else "🔴" if row[sentiment_col] < 0 else "🟡"

                st.markdown(f"""
                **{sentiment_color} {row['name']}** - {row['date'].strftime('%Y-%m-%d')} - ⭐ {row['star']:.0f} 星 - 情感: {row[sentiment_col]:.2f}

                **{selected_dimension}相關內容**:
                {row[reasons_col] if pd.notna(row[reasons_col]) else '（無相關評論）'}

                **完整評論**:
                {row['text']}

                ---
                """)

    else:
        st.warning(f"⚠️ 篩選後的數據中沒有 {selected_dimension} 的相關評論")

    st.markdown("---")

    # 評論瀏覽區
    st.markdown('<a id="reviews"></a>', unsafe_allow_html=True)
    st.subheader("💬 評論內容瀏覽")

    # 排序選項
    sort_option = st.selectbox(
        "排序方式",
        ["最新", "最舊", "最高分", "最低分"]
    )

    if sort_option == "最新":
        display_df = filtered_df.sort_values('date', ascending=False)
    elif sort_option == "最舊":
        display_df = filtered_df.sort_values('date', ascending=True)
    elif sort_option == "最高分":
        display_df = filtered_df.sort_values('star', ascending=False)
    else:
        display_df = filtered_df.sort_values('star', ascending=True)

    # 顯示評論
    show_columns = ['date', 'name', 'star', 'sentiment', 'text']
    display_data = display_df[show_columns].head(10).copy()
    display_data['sentiment'] = display_data['sentiment'].map(sentiment_map)
    display_data.columns = ['日期', '姓名', '星級', '情感', '評論內容']

    st.dataframe(display_data, use_container_width=True, height=400)

    # 下載功能
    st.markdown("---")
    st.markdown('<a id="download"></a>', unsafe_allow_html=True)
    st.subheader("📥 資料下載")

    csv = filtered_df.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        label="下載篩選後的資料 (CSV)",
        data=csv,
        file_name=f"w_hotel_reviews_filtered_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv"
    )

except Exception as e:
    st.error(f"發生錯誤: {str(e)}")
    st.info("請確保 'chat_W_hotel.xlsx' 檔案在相同目錄下")
