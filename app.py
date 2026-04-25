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
            
            # Formato DD.MM.YY
            if isinstance(date_str, str) and '.' in date_str:
                parts = date_str.split('.')
                if len(parts) == 3:
                    day, month, year = parts
                    year = f"20{year}" if len(year) == 2 else year
                    return datetime(int(year), int(month), int(day))
            
            # Intentar parsing automático
            return pd.to_datetime(date_str)
        except:
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
                balance_col = next((col for col in data_df.columns if str(col).strip().lower() == 'balance'), None)
                explanation_col = next((col for col in data_df.columns if str(col).strip().lower() == 'explanation'), None)
                
                if not date_col:
                    continue
                
                # Limpiar y procesar datos
                data_df = data_df[pd.notna(data_df[date_col])]
                
                for _, row in data_df.iterrows():
                    transaction = {
                        'Date': self.parse_date(row[date_col]) if date_col else None,
                        'Payer': row[payer_col] if payer_col else '',
                        'Recipient': row[recipient_col] if recipient_col else '',
                        'Transaction': row[transaction_col] if transaction_col else '',
                        'Out': pd.to_numeric(row[out_col], errors='coerce') if out_col else 0,
                        'In': pd.to_numeric(row[in_col], errors='coerce') if in_col else 0,
                        'Balance': pd.to_numeric(row[balance_col], errors='coerce') if balance_col else 0,
                        'Explanation': row[explanation_col] if explanation_col else '',
                        'Source_File': filename,
                        'Source_Sheet': sheet_name
                    }
                    
                    # Solo agregar si tiene fecha válida
                    if transaction['Date']:
                        all_transactions.append(transaction)
            
            return pd.DataFrame(all_transactions)
            
        except Exception as e:
            st.error(f"Error al procesar {filename}: {str(e)}")
            return None
    
    def consolidate_reports(self, uploaded_files):
        """Consolida múltiples reportes financieros"""
        all_data = []
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for idx, uploaded_file in enumerate(uploaded_files):
            status_text.text(f"Procesando {uploaded_file.name}...")
            
            df = self.load_financial_report(uploaded_file, uploaded_file.name)
            if df is not None and not df.empty:
                all_data.append(df)
            
            progress_bar.progress((idx + 1) / len(uploaded_files))
        
        status_text.empty()
        progress_bar.empty()
        
        if all_data:
            self.transactions_df = pd.concat(all_data, ignore_index=True)
            self.transactions_df = self.transactions_df.sort_values('Date')
            self.transactions_df['Year'] = self.transactions_df['Date'].dt.year
            self.transactions_df['Month'] = self.transactions_df['Date'].dt.month
            self.transactions_df['YearMonth'] = self.transactions_df['Date'].dt.to_period('M')
            
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
            
            # Hoja 2: Resumen mensual
            monthly_summary = self.transactions_df.groupby('YearMonth').agg({
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
    st.markdown("**Consolida y analiza tus reportes financieros mensuales de manera eficiente**")
    
    # Inicializar el consolidador
    if 'consolidator' not in st.session_state:
        st.session_state.consolidator = FinancialConsolidator()
    
    consolidator = st.session_state.consolidator
    
    # Sidebar
    with st.sidebar:
        st.header("⚙️ Configuración")
        
        uploaded_files = st.file_uploader(
            "Cargar reportes financieros",
            type=['xlsx', 'xls'],
            accept_multiple_files=True,
            help="Selecciona uno o más archivos Excel con reportes financieros"
        )
        
        if uploaded_files:
            if st.button("🔄 Consolidar Reportes", type="primary", use_container_width=True):
                with st.spinner("Consolidando reportes..."):
                    if consolidator.consolidate_reports(uploaded_files):
                        st.success(f"✅ {len(uploaded_files)} archivos consolidados exitosamente")
                        st.session_state.consolidated = True
                    else:
                        st.error("❌ Error al consolidar reportes")
        
        st.divider()
        
        if hasattr(st.session_state, 'consolidated') and st.session_state.consolidated:
            st.subheader("📥 Exportar Datos")
            
            excel_data = consolidator.export_to_excel()
            st.download_button(
                label="⬇️ Descargar Excel Consolidado",
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
                label="💰 Total Ingresos",
                value=f"${stats['total_income']:,.2f}",
                delta=None
            )
        
        with col2:
            st.metric(
                label="💸 Total Gastos",
                value=f"${stats['total_expenses']:,.2f}",
                delta=None
            )
        
        with col3:
            st.metric(
                label="📊 Flujo Neto",
                value=f"${stats['net_flow']:,.2f}",
                delta=None,
                delta_color="normal" if stats['net_flow'] >= 0 else "inverse"
            )
        
        with col4:
            st.metric(
                label="🔢 Transacciones",
                value=f"{stats['total_transactions']:,}",
                delta=None
            )
        
        st.divider()
        
        # Tabs para diferentes vistas
        tab1, tab2, tab3, tab4 = st.tabs(["📈 Análisis", "📋 Transacciones", "📊 Categorías", "👥 Personas"])
        
        with tab1:
            st.subheader("Análisis de Flujo de Caja")
            
            # Gráfico de flujo mensual
            monthly_data = consolidator.transactions_df.groupby('YearMonth').agg({
                'In': 'sum',
                'Out': 'sum'
            }).reset_index()
            monthly_data['YearMonth'] = monthly_data['YearMonth'].astype(str)
            monthly_data['Net'] = monthly_data['In'] - monthly_data['Out']
            
            fig = go.Figure()
            fig.add_trace(go.Bar(
                name='Ingresos',
                x=monthly_data['YearMonth'],
                y=monthly_data['In'],
                marker_color='#28a745'
            ))
            fig.add_trace(go.Bar(
                name='Gastos',
                x=monthly_data['YearMonth'],
                y=monthly_data['Out'],
                marker_color='#dc3545'
            ))
            fig.add_trace(go.Scatter(
                name='Flujo Neto',
                x=monthly_data['YearMonth'],
                y=monthly_data['Net'],
                mode='lines+markers',
                line=dict(color='#007bff', width=3),
                marker=dict(size=8)
            ))
            
            fig.update_layout(
                title='Flujo de Caja Mensual',
                xaxis_title='Mes',
                yaxis_title='Monto ($)',
                barmode='group',
                height=500,
                hovermode='x unified'
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            # Balance acumulado
            col1, col2 = st.columns(2)
            
            with col1:
                # Top gastos
                top_expenses = consolidator.transactions_df.groupby('Transaction')['Out'].sum().sort_values(ascending=False).head(10)
                fig_expenses = px.bar(
                    x=top_expenses.values,
                    y=top_expenses.index,
                    orientation='h',
                    title='Top 10 Categorías de Gastos',
                    labels={'x': 'Monto ($)', 'y': 'Categoría'},
                    color=top_expenses.values,
                    color_continuous_scale='Reds'
                )
                fig_expenses.update_layout(showlegend=False, height=400)
                st.plotly_chart(fig_expenses, use_container_width=True)
            
            with col2:
                # Distribución por persona
                payer_dist = consolidator.transactions_df.groupby('Payer')['Out'].sum().sort_values(ascending=False)
                fig_payer = px.pie(
                    values=payer_dist.values,
                    names=payer_dist.index,
                    title='Distribución de Gastos por Persona',
                    hole=0.4
                )
                fig_payer.update_layout(height=400)
                st.plotly_chart(fig_payer, use_container_width=True)
        
        with tab2:
            st.subheader("Todas las Transacciones")
            
            # Filtros
            col1, col2, col3 = st.columns(3)
            
            with col1:
                years = sorted(consolidator.transactions_df['Year'].unique())
                selected_year = st.selectbox("Año", ['Todos'] + list(years))
            
            with col2:
                months = sorted(consolidator.transactions_df['Month'].unique())
                selected_month = st.selectbox("Mes", ['Todos'] + list(months))
            
            with col3:
                categories = sorted(consolidator.transactions_df['Transaction'].unique())
                selected_category = st.selectbox("Categoría", ['Todas'] + list(categories))
            
            # Aplicar filtros
            filtered_df = consolidator.transactions_df.copy()
            
            if selected_year != 'Todos':
                filtered_df = filtered_df[filtered_df['Year'] == selected_year]
            
            if selected_month != 'Todos':
                filtered_df = filtered_df[filtered_df['Month'] == selected_month]
            
            if selected_category != 'Todas':
                filtered_df = filtered_df[filtered_df['Transaction'] == selected_category]
            
            # Mostrar tabla
            display_df = filtered_df[['Date', 'Payer', 'Recipient', 'Transaction', 'Out', 'In', 'Balance', 'Explanation']].copy()
            display_df['Date'] = display_df['Date'].dt.strftime('%Y-%m-%d')
            
            st.dataframe(
                display_df,
                use_container_width=True,
                height=500,
                hide_index=True
            )
            
            st.info(f"Mostrando {len(filtered_df)} de {len(consolidator.transactions_df)} transacciones")
        
        with tab3:
            st.subheader("Análisis por Categorías")
            
            category_summary = consolidator.transactions_df.groupby('Transaction').agg({
                'Out': ['sum', 'mean', 'count'],
                'In': 'sum'
            }).round(2)
            
            category_summary.columns = ['Total Gastos', 'Promedio', 'Cantidad', 'Total Ingresos']
            category_summary = category_summary.sort_values('Total Gastos', ascending=False)
            
            st.dataframe(category_summary, use_container_width=True)
        
        with tab4:
            st.subheader("Análisis por Personas")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("#### Pagadores (Payer)")
                payer_summary = consolidator.transactions_df.groupby('Payer').agg({
                    'Out': ['sum', 'count'],
                    'In': 'sum'
                }).round(2)
                payer_summary.columns = ['Total Pagado', 'Transacciones', 'Total Recibido']
                payer_summary = payer_summary.sort_values('Total Pagado', ascending=False)
                st.dataframe(payer_summary, use_container_width=True)
            
            with col2:
                st.markdown("#### Receptores (Recipient)")
                recipient_summary = consolidator.transactions_df.groupby('Recipient').agg({
                    'In': ['sum', 'count'],
                    'Out': 'sum'
                }).round(2)
                recipient_summary.columns = ['Total Recibido', 'Transacciones', 'Total Pagado']
                recipient_summary = recipient_summary.sort_values('Total Recibido', ascending=False)
                st.dataframe(recipient_summary, use_container_width=True)
    
    else:
        # Estado inicial
        st.info("👆 Carga tus archivos Excel en el panel lateral para comenzar")
        
        st.markdown("""
        ### 🚀 Cómo usar esta aplicación:
        
        1. **Carga tus archivos**: Usa el botón en el panel lateral para cargar uno o más reportes financieros en formato Excel
        2. **Consolida**: Haz clic en "Consolidar Reportes" para procesar los archivos
        3. **Analiza**: Explora los diferentes tabs para ver análisis detallados
        4. **Exporta**: Descarga el reporte consolidado en formato Excel
        
        ### 📋 Formatos soportados:
        - Archivos .xlsx y .xls
        - Estructura estándar de reportes de Gojitech
        - Múltiples hojas por archivo
        
        ### 💡 Características:
        - ✅ Consolidación automática de múltiples archivos
        - ✅ Análisis de flujo de caja mensual
        - ✅ Gráficos interactivos
        - ✅ Filtros dinámicos
        - ✅ Exportación a Excel con múltiples hojas
        - ✅ Análisis por categorías y personas
        """)


if __name__ == "__main__":
    main()
