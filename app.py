import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io
import re
from pathlib import Path

st.set_page_config(
    page_title="Financial Report Consolidator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        color: #1f1f1f;
        margin-bottom: 1rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 10px;
        color: white;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .success-msg {
        padding: 1rem;
        background-color: #d4edda;
        border-left: 4px solid #28a745;
        border-radius: 4px;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)


class FinancialConsolidator:
    """Clase principal para consolidar reportes financieros"""
    
    def __init__(self):
        self.transactions_df = None
        self.metadata = {}
        
    def extract_date_from_filename(self, filename):
        """Extrae mes y año del nombre del archivo"""
        months_es = {
            'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4,
            'mayo': 5, 'junio': 6, 'julio': 7, 'agosto': 8,
            'septiembre': 9, 'octubre': 10, 'noviembre': 11, 'diciembre': 12
        }
        months_en = {
            'january': 1, 'february': 2, 'march': 3, 'april': 4,
            'may': 5, 'june': 6, 'july': 7, 'august': 8,
            'september': 9, 'october': 10, 'november': 11, 'december': 12,
            'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'jun': 6,
            'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
        }
        
        filename_lower = filename.lower()
        
        # Buscar año (4 dígitos)
        year_match = re.search(r'20\d{2}', filename_lower)
        year = int(year_match.group()) if year_match else None
        
        # Buscar mes
        month = None
        for month_name, month_num in {**months_es, **months_en}.items():
            if month_name in filename_lower:
                month = month_num
                break
        
        return month, year
    
    def parse_date(self, date_str):
        """Convierte fecha de diferentes formatos a datetime"""
        if pd.isna(date_str):
            return None
            
        try:
            # Si ya es datetime
            if isinstance(date_str, datetime):
                return date_str
            
            # Convertir a string y limpiar espacios
            date_str = str(date_str).strip().replace(' ', '')
            
            # Formato DD.MM.YY o DD.MM.YYYY
            # Ejemplos: "6.3.24", "30.4.25", "4.3 24" (con espacios), "12.03.2024"
            if '.' in date_str:
                parts = date_str.split('.')
                
                if len(parts) == 3:
                    day = parts[0]
                    month = parts[1]
                    year = parts[2]
                    
                    # Convertir año de 2 dígitos a 4 dígitos
                    if len(year) <= 2:
                        year_int = int(year)
                        # Si es 00-49, asumir 2000-2049
                        # Si es 50-99, asumir 1950-1999
                        if year_int < 50:
                            year = f"20{year.zfill(2)}"
                        else:
                            year = f"19{year.zfill(2)}"
                    
                    return datetime(int(year), int(month), int(day))
            
            # Intentar parsing automático de pandas
            parsed = pd.to_datetime(date_str, dayfirst=True)
            return parsed
            
        except Exception as e:
            # Si falla, retornar None
            return None
    
    def load_financial_report(self, file, filename):
        """Carga un reporte financiero desde un archivo Excel"""
        try:
            # Leer todas las hojas
            xl = pd.ExcelFile(file)
            
            all_transactions = []
            
            for sheet_name in xl.sheet_names:
                # Leer sin encabezados
                df = pd.read_excel(file, sheet_name=sheet_name, header=None)
                
                # Buscar la fila con "Date", "Payer", "Recipient", etc.
                header_row = None
                for idx, row in df.iterrows():
                    if any(str(cell).strip().lower() == 'date' for cell in row if pd.notna(cell)):
                        header_row = idx
                        break
                
                if header_row is None:
                    continue
                
                # Extraer columnas relevantes
                headers = df.iloc[header_row].tolist()
                data_df = df.iloc[header_row + 1:].copy()
                data_df.columns = headers
                
                # Identificar columnas clave
                date_col = next((col for col in data_df.columns if str(col).strip().lower() == 'date'), None)
                payer_col = next((col for col in data_df.columns if str(col).strip().lower() == 'payer'), None)
                recipient_col = next((col for col in data_df.columns if str(col).strip().lower() == 'recipient'), None)
                transaction_col = next((col for col in data_df.columns if str(col).strip().lower() == 'transaction'), None)
                out_col = next((col for col in data_df.columns if str(col).strip().lower() == 'out'), None)
                in_col = next((col for col in data_df.columns if str(col).strip().lower() == 'in'), None)
                explanation_col = next((col for col in data_df.columns if str(col).strip().lower() == 'explanation'), None)
                
                # NO incluir Balance
                
                if not date_col:
                    continue
                
                # INCLUIR SOLO FILAS QUE TENGAN VALOR EN OUT O IN
                for _, row in data_df.iterrows():
                    # Obtener valores de Out e In
                    out_value = pd.to_numeric(row[out_col], errors='coerce') if out_col else 0
                    in_value = pd.to_numeric(row[in_col], errors='coerce') if in_col else 0
                    
                    # Saltar filas donde AMBOS Out e In son 0 o NaN
                    if pd.isna(out_value) or out_value == 0:
                        if pd.isna(in_value) or in_value == 0:
                            continue  # Saltar esta fila
                    
                    # Parsear fecha (puede ser None)
                    parsed_date = self.parse_date(row[date_col]) if date_col else None
                    
                    transaction = {
                        'Date': parsed_date,
                        'Payer': row[payer_col] if payer_col and pd.notna(row[payer_col]) else '',
                        'Recipient': row[recipient_col] if recipient_col and pd.notna(row[recipient_col]) else '',
                        'Transaction': row[transaction_col] if transaction_col and pd.notna(row[transaction_col]) else '',
                        'Out': out_value if pd.notna(out_value) else 0,
                        'In': in_value if pd.notna(in_value) else 0,
                        'Explanation': row[explanation_col] if explanation_col and pd.notna(row[explanation_col]) else '',
                        'Source_File': filename
                    }
                    
                    # Agregar la transacción
                    all_transactions.append(transaction)
            
            return pd.DataFrame(all_transactions)
            
        except Exception as e:
            st.error(f"Error processing {filename}: {str(e)}")
            return None
    
    def consolidate_reports(self, uploaded_files):
        """Consolida múltiples reportes financieros"""
        all_data = []
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for idx, uploaded_file in enumerate(uploaded_files):
            status_text.text(f"Processing {uploaded_file.name}...")
            
            df = self.load_financial_report(uploaded_file, uploaded_file.name)
            if df is not None and not df.empty:
                all_data.append(df)
            
            progress_bar.progress((idx + 1) / len(uploaded_files))
        
        status_text.empty()
        progress_bar.empty()
        
        if all_data:
            self.transactions_df = pd.concat(all_data, ignore_index=True)
            
            # Ordenar por fecha solo las que tienen fecha válida
            self.transactions_df = self.transactions_df.sort_values('Date', na_position='last')
            
            # Crear columnas Year y Month (manejar fechas None)
            self.transactions_df['Year'] = self.transactions_df['Date'].apply(
                lambda x: x.year if pd.notna(x) else None
            )
            self.transactions_df['Month'] = self.transactions_df['Date'].apply(
                lambda x: x.month if pd.notna(x) else None
            )
            self.transactions_df['YearMonth'] = self.transactions_df['Date'].apply(
                lambda x: x.to_period('M') if pd.notna(x) else None
            )
            
            return True
        return False
    
    def get_summary_stats(self):
        """Genera estadísticas resumidas"""
        if self.transactions_df is None or self.transactions_df.empty:
            return {}
        
        df = self.transactions_df
        
        return {
            'total_transactions': len(df),
            'total_income': df['In'].sum(),
            'total_expenses': df['Out'].sum(),
            'net_flow': df['In'].sum() - df['Out'].sum(),
            'date_range': f"{df['Date'].min().strftime('%Y-%m-%d')} to {df['Date'].max().strftime('%Y-%m-%d')}",
            'unique_payers': df['Payer'].nunique(),
            'unique_recipients': df['Recipient'].nunique(),
            'files_processed': df['Source_File'].nunique()
        }
    
    def export_to_excel(self):
        """Exporta datos consolidados a Excel con múltiples hojas"""
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Hoja 1: Todas las transacciones
            self.transactions_df.to_excel(writer, sheet_name='All Transactions', index=False)
            
            # Hoja 2: Resumen mensual (solo para filas con fecha válida)
            df_with_dates = self.transactions_df[self.transactions_df['YearMonth'].notna()].copy()
            if not df_with_dates.empty:
                monthly_summary = df_with_dates.groupby('YearMonth').agg({
                    'In': 'sum',
                    'Out': 'sum',
                    'Date': 'count'
                }).rename(columns={'Date': 'Transaction Count'})
                monthly_summary['Net Flow'] = monthly_summary['In'] - monthly_summary['Out']
                monthly_summary.to_excel(writer, sheet_name='Monthly Summary')
            
            # Hoja 3: Resumen por categoría
            category_summary = self.transactions_df.groupby('Transaction').agg({
                'Out': 'sum',
                'In': 'sum',
                'Date': 'count'
            }).rename(columns={'Date': 'Count'}).sort_values('Out', ascending=False)
            category_summary.to_excel(writer, sheet_name='By Category')
            
            # Hoja 4: Resumen por persona
            payer_summary = self.transactions_df.groupby('Payer').agg({
                'Out': 'sum',
                'Date': 'count'
            }).rename(columns={'Date': 'Count'}).sort_values('Out', ascending=False)
            payer_summary.to_excel(writer, sheet_name='By Payer')
            
        output.seek(0)
        return output


def main():
    st.markdown('<h1 class="main-header">📊 Financial Report Consolidator</h1>', unsafe_allow_html=True)
    st.markdown("**Consolidate and analyze your monthly financial reports efficiently**")
    
    # Inicializar el consolidador
    if 'consolidator' not in st.session_state:
        st.session_state.consolidator = FinancialConsolidator()
    
    consolidator = st.session_state.consolidator
    
    # Sidebar
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        uploaded_files = st.file_uploader(
            "Upload financial reports",
            type=['xlsx', 'xls'],
            accept_multiple_files=True,
            help="Select one or more Excel files with financial reports"
        )
        
        if uploaded_files:
            if st.button("🔄 Consolidate Reports", type="primary", use_container_width=True):
                with st.spinner("Consolidating reports..."):
                    if consolidator.consolidate_reports(uploaded_files):
                        st.success(f"✅ {len(uploaded_files)} files consolidated successfully")
                        st.session_state.consolidated = True
                    else:
                        st.error("❌ Error consolidating reports")
        
        st.divider()
        
        if hasattr(st.session_state, 'consolidated') and st.session_state.consolidated:
            st.subheader("📥 Export Data")
            
            excel_data = consolidator.export_to_excel()
            st.download_button(
                label="⬇️ Download Consolidated Excel",
                data=excel_data,
                file_name=f"financial_consolidated_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    # Main content
    if hasattr(st.session_state, 'consolidated') and st.session_state.consolidated:
        stats = consolidator.get_summary_stats()
        
        # Métricas principales
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric(
                label="💰 Total Income",
                value=f"${stats['total_income']:,.2f}",
                delta=None
            )
        
        with col2:
            st.metric(
                label="💸 Total Expenses",
                value=f"${stats['total_expenses']:,.2f}",
                delta=None
            )
        
        with col3:
            st.metric(
                label="📊 Net Flow",
                value=f"${stats['net_flow']:,.2f}",
                delta=None,
                delta_color="normal" if stats['net_flow'] >= 0 else "inverse"
            )
        
        with col4:
            st.metric(
                label="🔢 Transactions",
                value=f"{stats['total_transactions']:,}",
                delta=None
            )
        
        st.divider()
        
        # Tabs para diferentes vistas
        tab1, tab2, tab3, tab4 = st.tabs(["📈 Analysis", "📋 Transactions", "📊 Categories", "👥 People"])
        
        with tab1:
            st.subheader("Cash Flow Analysis")
            
            # Filtrar solo filas con fechas válidas para gráficos
            df_with_dates = consolidator.transactions_df[consolidator.transactions_df['YearMonth'].notna()].copy()
            
            if df_with_dates.empty:
                st.warning("No data with valid dates available for analysis.")
            else:
                # Gráfico de flujo mensual
                monthly_data = df_with_dates.groupby('YearMonth').agg({
                    'In': 'sum',
                    'Out': 'sum'
                }).reset_index()
                monthly_data['YearMonth'] = monthly_data['YearMonth'].astype(str)
                monthly_data['Net'] = monthly_data['In'] - monthly_data['Out']
            
                fig = go.Figure()
                fig.add_trace(go.Bar(
                    name='Income',
                    x=monthly_data['YearMonth'],
                    y=monthly_data['In'],
                    marker_color='#28a745'
                ))
                fig.add_trace(go.Bar(
                    name='Expenses',
                    x=monthly_data['YearMonth'],
                    y=monthly_data['Out'],
                    marker_color='#dc3545'
                ))
                fig.add_trace(go.Scatter(
                    name='Net Flow',
                    x=monthly_data['YearMonth'],
                    y=monthly_data['Net'],
                    mode='lines+markers',
                    line=dict(color='#007bff', width=3),
                    marker=dict(size=8)
                ))
                
                fig.update_layout(
                    title='Monthly Cash Flow',
                    xaxis_title='Month',
                    yaxis_title='Amount ($)',
                    barmode='group',
                    height=500,
                    hovermode='x unified'
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
                # Balance acumulado
                col1, col2 = st.columns(2)
                
                with col1:
                    # Top gastos
                    top_expenses = df_with_dates.groupby('Transaction')['Out'].sum().sort_values(ascending=False).head(10)
                    fig_expenses = px.bar(
                        x=top_expenses.values,
                        y=top_expenses.index,
                        orientation='h',
                        title='Top 10 Expense Categories',
                        labels={'x': 'Amount ($)', 'y': 'Category'},
                        color=top_expenses.values,
                        color_continuous_scale='Reds'
                    )
                    fig_expenses.update_layout(showlegend=False, height=400)
                    st.plotly_chart(fig_expenses, use_container_width=True)
                
                with col2:
                    # Distribución por persona
                    payer_dist = df_with_dates.groupby('Payer')['Out'].sum().sort_values(ascending=False)
                    fig_payer = px.pie(
                        values=payer_dist.values,
                        names=payer_dist.index,
                        title='Expense Distribution by Person',
                        hole=0.4
                    )
                    fig_payer.update_layout(height=400)
                    st.plotly_chart(fig_payer, use_container_width=True)
        
        with tab2:
            st.subheader("All Transactions")
            
            # Verificar que hay datos
            if consolidator.transactions_df is None or consolidator.transactions_df.empty:
                st.warning("No transaction data available. Please upload and consolidate files first.")
            else:
                # Filtros
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    years = sorted(consolidator.transactions_df['Year'].unique())
                    selected_year = st.selectbox("Year", ['All'] + list(years))
                
                with col2:
                    months = sorted(consolidator.transactions_df['Month'].unique())
                    selected_month = st.selectbox("Month", ['All'] + list(months))
                
                with col3:
                    categories = sorted(consolidator.transactions_df['Transaction'].unique())
                    selected_category = st.selectbox("Category", ['All'] + list(categories))
                
                # Aplicar filtros
                filtered_df = consolidator.transactions_df.copy()
                
                if selected_year != 'All':
                    filtered_df = filtered_df[filtered_df['Year'] == selected_year]
                
                if selected_month != 'All':
                    filtered_df = filtered_df[filtered_df['Month'] == selected_month]
                
                if selected_category != 'All':
                    filtered_df = filtered_df[filtered_df['Transaction'] == selected_category]
                
                # Mostrar tabla
                display_df = filtered_df[['Date', 'Payer', 'Recipient', 'Transaction', 'Out', 'In', 'Explanation']].copy()
                display_df['Date'] = display_df['Date'].dt.strftime('%Y-%m-%d')
                
                st.dataframe(
                    display_df,
                    use_container_width=True,
                    height=500,
                    hide_index=True
                )
                
                st.info(f"Showing {len(filtered_df)} of {len(consolidator.transactions_df)} transactions")
        
        with tab3:
            st.subheader("Analysis by Categories")
            
            if consolidator.transactions_df is None or consolidator.transactions_df.empty:
                st.warning("No transaction data available. Please upload and consolidate files first.")
            else:
                category_summary = consolidator.transactions_df.groupby('Transaction').agg({
                    'Out': ['sum', 'mean', 'count'],
                    'In': 'sum'
                }).round(2)
                
                category_summary.columns = ['Total Expenses', 'Average', 'Quantity', 'Total Income']
                category_summary = category_summary.sort_values('Total Expenses', ascending=False)
                
                st.dataframe(category_summary, use_container_width=True)
        
        with tab4:
            st.subheader("Analysis by People")
            
            if consolidator.transactions_df is None or consolidator.transactions_df.empty:
                st.warning("No transaction data available. Please upload and consolidate files first.")
            else:
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("#### Payers")
                    payer_summary = consolidator.transactions_df.groupby('Payer').agg({
                        'Out': ['sum', 'count'],
                        'In': 'sum'
                    }).round(2)
                    payer_summary.columns = ['Total Paid', 'Transactions', 'Total Received']
                    payer_summary = payer_summary.sort_values('Total Paid', ascending=False)
                    st.dataframe(payer_summary, use_container_width=True)
                
                with col2:
                    st.markdown("#### Recipients")
                    recipient_summary = consolidator.transactions_df.groupby('Recipient').agg({
                        'In': ['sum', 'count'],
                        'Out': 'sum'
                    }).round(2)
                    recipient_summary.columns = ['Total Received', 'Transactions', 'Total Paid']
                    recipient_summary = recipient_summary.sort_values('Total Received', ascending=False)
                    st.dataframe(recipient_summary, use_container_width=True)
    
    else:
        # Estado inicial
        st.info("👆 Upload your Excel files in the sidebar to get started")
        
        st.markdown("""
        ### 🚀 How to use this app:
        
        1. **Upload your files**: Use the button in the sidebar to upload one or more financial reports in Excel format
        2. **Consolidate**: Click "Consolidate Reports" to process the files
        3. **Analyze**: Explore the different tabs to see detailed analysis
        4. **Export**: Download the consolidated report in Excel format
        
        ### 📋 Supported formats:
        - .xlsx and .xls files
        - Standard financial report structure
        - Multiple sheets per file
        
        ### 💡 Features:
        - ✅ Automatic consolidation of multiple files
        - ✅ Monthly cash flow analysis
        - ✅ Interactive charts
        - ✅ Dynamic filters
        - ✅ Excel export with multiple sheets
        - ✅ Analysis by categories and people
        """)


if __name__ == "__main__":
    main()
