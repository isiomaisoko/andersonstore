import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(layout='wide', page_title='Anderson E-commerce Sales Dashboard')
st.title('Anderson Store Sales Performance & Analytics Dashboard')

# File upload
uploaded_file = st.file_uploader("Upload E-commerce Excel File", type=['xlsx'])

if uploaded_file:
    @st.cache_data
    def load_data(file):
        return pd.read_excel(file, sheet_name='Data')

    df = load_data(uploaded_file)
    df['OrderDate'] = pd.to_datetime(df['OrderDate'])

    # =============================
    # Sidebar filters
    # =============================
    with st.sidebar:
        st.header('Filters')
        products = st.multiselect('Product', df['Product'].unique(), default=df['Product'].unique())
        categories = st.multiselect('Category', df['Category'].unique(), default=df['Category'].unique())
        regions = st.multiselect('Region', df['Region'].unique(), default=df['Region'].unique())
        date_min = st.date_input('Start Date', df['OrderDate'].min())
        date_max = st.date_input('End Date', df['OrderDate'].max())

    # Apply filters
    mask = (
        df['Product'].isin(products) &
        df['Category'].isin(categories) &
        df['Region'].isin(regions) &
        (df['OrderDate'] >= pd.to_datetime(date_min)) &
        (df['OrderDate'] <= pd.to_datetime(date_max))
    )
    df_filtered = df.loc[mask]

    # =============================
    # DYNAMIC TOP KPIs
    # =============================
    col1, col2, col3, col4, col5 = st.columns(5)
    total_revenue = df_filtered['Revenue'].sum()
    total_units = df_filtered['UnitsSold'].sum()
    total_profit = df_filtered['Profit'].sum()
    unique_orders = df_filtered['OrderID'].nunique() if df_filtered['OrderID'].nunique() > 0 else 1
    avg_order_value = total_revenue / unique_orders
    avg_profit_order = total_profit / unique_orders

    col1.metric('Total Revenue', f"${total_revenue:,.2f}")
    col2.metric('Total Units Sold', f"{int(total_units):,}")
    col3.metric('Total Profit', f"${total_profit:,.2f}")
    col4.metric('Average Order Value', f"${avg_order_value:,.2f}")
    col5.metric('Avg Profit per Order', f"${avg_profit_order:,.2f}")

    # Download filtered CSV
    csv = df_filtered.to_csv(index=False).encode('utf-8')
    st.download_button("Download Filtered Data", csv, "filtered_sales_data.csv", "text/csv")

    # =============================
    # DYNAMIC CHARTS
    # =============================
    # Revenue & Profit over time
    if not df_filtered.empty:
        rev_time = df_filtered.groupby(df_filtered['OrderDate'].dt.to_period('M')).agg({'Revenue':'sum', 'Profit':'sum'}).reset_index()
        rev_time['OrderDate'] = rev_time['OrderDate'].dt.to_timestamp()

        fig1 = px.line(rev_time, x='OrderDate', y='Revenue', markers=True, title='Revenue by Month')
        fig2 = px.line(rev_time, x='OrderDate', y='Profit', markers=True, title='Profit by Month')
        col1, col2 = st.columns(2)
        col1.plotly_chart(fig1, use_container_width=True)
        col2.plotly_chart(fig2, use_container_width=True)

        # Revenue by Product
        prod_rev = df_filtered.groupby('Product')['Revenue'].sum().reset_index().sort_values('Revenue', ascending=False)
        fig3 = px.bar(prod_rev, x='Product', y='Revenue', title='Revenue by Product', text_auto=True)
        st.plotly_chart(fig3, use_container_width=True)

        # Revenue by Category (Pie Chart)
        cat_rev = df_filtered.groupby('Category')['Revenue'].sum().reset_index()
        fig4 = px.pie(cat_rev, names='Category', values='Revenue', title='Revenue by Category', hole=0.3)
        st.plotly_chart(fig4, use_container_width=True)

    # =============================
    # SUMMARY REPORT & EXECUTIVE INSIGHTS
    # =============================
    st.markdown("## 📊 Sales Summary Report")
    top_product = prod_rev.iloc[0]['Product'] if not prod_rev.empty else "N/A"
    top_product_revenue = prod_rev.iloc[0]['Revenue'] if not prod_rev.empty else 0
    top_category = cat_rev.iloc[cat_rev['Revenue'].idxmax()]['Category'] if not cat_rev.empty else "N/A"
    top_category_revenue = cat_rev['Revenue'].max() if not cat_rev.empty else 0
    region_rev = df_filtered.groupby('Region')['Revenue'].sum().reset_index()
    top_region = region_rev.iloc[region_rev['Revenue'].idxmax()]['Region'] if not region_rev.empty else "N/A"
    top_region_revenue = region_rev['Revenue'].max() if not region_rev.empty else 0
    trend = "increasing 📈" if len(df_filtered) > 1 and df_filtered['Revenue'].iloc[-1] > df_filtered['Revenue'].iloc[0] else "decreasing 📉"

    st.markdown(f"""
    **Overall Performance**

    - Total Revenue: **${total_revenue:,.2f}**
    - Total Units Sold: **{int(total_units):,} units**
    - Total Profit: **${total_profit:,.2f}**
    - Average Order Value: **${avg_order_value:,.2f}**
    - Average Profit per Order: **${avg_profit_order:,.2f}**

    **Top Performers**

    - Best Selling Product: **{top_product}** (Revenue: **${top_product_revenue:,.2f}**)
    - Top Revenue Category: **{top_category}** (Revenue: **${top_category_revenue:,.2f}**)
    - Highest Performing Region: **{top_region}** (Revenue: **${top_region_revenue:,.2f}**)

    **Trend Overview**

    - Revenue trend over time is **{trend}**.
    """)

    # Insights
    st.markdown("##  Executive Business Insights")
    insights = []
    if len(df_filtered) > 1 and df_filtered['Revenue'].iloc[-1] > df_filtered['Revenue'].iloc[0]:
        insights.append("Revenue is showing an upward trend, indicating positive business growth.")
    else:
        insights.append("Revenue is trending downward, suggesting potential issues.")

    profit_margin = (total_profit / total_revenue) * 100 if total_revenue > 0 else 0
    if profit_margin > 25:
        insights.append(f"Profit margins are strong at **{profit_margin:.1f}%**.")
    elif 10 <= profit_margin <= 25:
        insights.append(f"Profit margins are moderate at **{profit_margin:.1f}%**.")
    else:
        insights.append(f"Profit margins are low at **{profit_margin:.1f}%**.")

    top_product_share = (top_product_revenue / total_revenue) * 100 if total_revenue > 0 else 0
    insights.append(f"The product **{top_product}** contributes **{top_product_share:.1f}%** of revenue.")

    top_category_share = (top_category_revenue / total_revenue) * 100 if total_revenue > 0 else 0
    if top_category_share > 45:
        insights.append(f"The category **{top_category}** dominates sales with **{top_category_share:.1f}%**.")
    else:
        insights.append(f"Revenue is reasonably diversified, with **{top_category}** leading at **{top_category_share:.1f}%**.")

    insights.append(f"**{top_region}** is the highest-performing region.")

    for point in insights:
        st.markdown(f"- {point}")
else:
    st.warning("Please upload an Excel file to view the dashboard.")