import streamlit as st
import pandas as pd
from datetime import datetime
import os

# Настройка страницы
st.set_page_config(page_title="Учет заказов: Пошив штор", page_icon="🪡", layout="wide")

PRICE_FILE = "tarifs.xlsx"
MATERIALS_FILE = "materials.xlsx"
PRODUCTS_FILE = "products.xlsx"

# 1. Справочник цен на РАБОТЫ (услуги швей)
def init_default_price():
    default_data = {
        "Наименование работы": [
            "Пошив тюля с утяжелителем на шторной ленте", "Стянуть ленту под размер",
            "Пошив портьеры на подкладке", "Пошив портьеры на подкладке из бархата", "Надставка сверху и снизу",
            "Пошив римской шторы без подкладки на тесьме", "Изготовление и притачивание канта", "Евроуголок",
            "Чехол на молнии из бархата", "Изготовление и притачивание рулика"
        ],
        "Ед.изм": ["м.п.", "м.п.", "м.п.", "м.п.", "м.п.", "кв.м.", "м.п.", "шт.", "шт.", "м.п."],
        "Цена (руб)": [420, 150, 770, 1078, 600, 1450, 210, 350, 550, 250],
        "Наценка (%)": [0, 0, 0, 40, 0, 0, 0, 0, 0, 0],
        "Категория": ["Тюль", "Тюль", "Портьеры на подкладке", "Портьеры на подкладке", "Портьеры на подкладке", 
                      "Римская штора", "Римская штора", "Римская штора", "Декоративная подушка", "Декоративная подушка"]
    }
    df = pd.DataFrame(default_data)
    df.to_excel(PRICE_FILE, index=False)
    return df

# 2. Справочник цен на МАТЕРИАЛЫ И ФУРНИТУРУ
def init_default_materials():
    default_data = {
        "Наименование материала": [
            "Шторная лента 8 см", "Утяжелитель нижний",
            "Подкладка стандарт", "Шторная лента для портьер люкс",
            "Тесьма для римских штор", "Липучка велкро (комплект)", "Кольца прозрачные",
            "Внутренняя подушка (наполнитель)", "Молния потайная"
        ],
        "Ед.изм": ["м.п.", "м.п.", "м.п.", "м.п.", "м.п.", "м.п.", "шт.", "шт.", "шт."],
        "Цена (руб)": [150, 250, 350, 450, 80, 120, 15, 300, 60],
        "Категория": ["Тюль", "Тюль", "Портьеры на подкладке", "Портьеры на подкладке", 
                      "Римская штора", "Римская штора", "Римская штора", "Декоративная подушка", "Декоративная подушка"]
    }
    df = pd.DataFrame(default_data)
    df.to_excel(MATERIALS_FILE, index=False)
    return df

# 3. Справочник НАИМЕНОВАНИЙ ИЗДЕЛИЙ
def init_default_products():
    default_data = {
        "Наименование изделия": ["Римская штора", "Портьеры на подкладке", "Тюль", "Декоративная подушка"]
    }
    df = pd.DataFrame(default_data)
    df.to_excel(PRODUCTS_FILE, index=False)
    return df

# Загрузка и автоматическая проверка файлов
if os.path.exists(PRICE_FILE):
    try: df_price = pd.read_excel(PRICE_FILE)
    except: df_price = init_default_price()
else: df_price = init_default_price()

if os.path.exists(MATERIALS_FILE):
    try: df_materials = pd.read_excel(MATERIALS_FILE)
    except: df_materials = init_default_materials()
else: df_materials = init_default_materials()

if os.path.exists(PRODUCTS_FILE):
    try: df_products = pd.read_excel(PRODUCTS_FILE)
    except: df_products = init_default_products()
else: df_products = init_default_products()

PRODUCTS_LIST = df_products["Наименование изделия"].dropna().tolist()

# --- ИМИТАЦИЯ БАЗЫ ЗАКАЗОВ ---
if "orders_db" not in st.session_state: st.session_state.orders_db = {}
if "current_items" not in st.session_state: st.session_state.current_items = []

init_client_name, init_client_phone, init_designer, init_admin, init_project_name = "", "", "", "", ""
init_date = datetime.now().date()

# =========================================================
# 🗂️ ЛЕВАЯ ПАНЕЛЬ (САЙДБАР)
# =========================================================
with st.sidebar:
    st.markdown("<h2 style='margin: 0;'>ЗАКАЗ №</h2>", unsafe_allow_html=True)
    num_col1, num_col2 = st.columns(2)
    with num_col1:
        order_digits = st.number_input("Номер", min_value=1, max_value=10000, value=101, step=1, label_visibility="collapsed", key="order_num_input")
    with num_col2:
        st.text_input("Год", value="- 2026", disabled=True, label_visibility="collapsed", key="year_input")
        
    full_order_id = f"{order_digits}-2026"
    order_status = st.radio("Статус документа", ["Новый", "Существующий"], key="status_input")

    st.markdown("---")
    st.markdown("**Дата**")
    order_date = st.date_input("Выбор даты", value=init_date, key="date_input")

    st.markdown("**Название заказа / Проект**")
    project_name = st.text_input("Проект", value=init_project_name, placeholder="Например: Проект Роксана", key="project_input")

    st.markdown("**Заказчик**")
    client_name = st.text_input("ФИО", value=init_client_name, placeholder="ФИО заказчика", key="client_input")
    st.markdown("**Контакты заказчика**")
    client_phone = st.text_input("Телефон", value=init_client_phone, placeholder="+7...", key="phone_input")

    st.markdown("---")
    st.markdown("**Дизайнер**")
    designer_name = st.text_input("Дизайнер проекта", value=init_designer, placeholder="Укажите дизайнера", key="designer_input")
    st.markdown("**Администратор**")
    admin_name = st.text_input("Менеджер/Админ", value=init_admin, placeholder="Кто принял заказ", key="admin_input")
# =========================================================
# 🖥️ ОСНОВНОЕ ОКНО
# =========================================================
display_project_title = f" — «{project_name}»" if project_name else ""
st.markdown(f"## 🛠️ Расчет заказа № {full_order_id}{display_project_title}")

st.subheader("➕ Добавление позиции в расчет")
with st.container(border=True):
    col_group1, col_group2, col_group3 = st.columns(3)
    
    with col_group1:
        product_options = ["[ + Создать новое изделие в справочник ]"] + PRODUCTS_LIST
        selected_product = st.selectbox("📌 Наименование изделия:", product_options, index=1 if len(product_options) > 1 else 0, key="prod_select")
        item_group = selected_product

    with col_group2:
        entry_type = st.radio("⚡ Тип позиции:", ["🧵 Работа", "📦 Материал"], horizontal=True)

    with col_group3:
        if entry_type == "🧵 Работа":
            df_filtered = df_price[df_price["Категория"] == item_group]
            filtered_dict = df_filtered.set_index("Наименование работы").to_dict(orient="index")
            select_ops = ["[ + Добавить новую РАБОТУ в Excel ]"] + list(filtered_dict.keys())
            selected_item = st.selectbox("Выберите наименование работы:", select_ops, key="work_select")
        else:
            df_filtered = df_materials[df_materials["Категория"] == item_group]
            filtered_dict = df_filtered.set_index("Наименование материала").to_dict(orient="index")
            select_ops = ["[ + Добавить новый МАТЕРИАЛ в Excel ]"] + list(filtered_dict.keys())
            selected_item = st.selectbox("Выберите наименование материала/фурнитуры:", select_ops, key="mat_select")

    if selected_item in ["[ + Добавить новую РАБОТУ в Excel ]", "[ + Добавить новый МАТЕРИАЛ в Excel ]"]:
        st.markdown(f"📎 *Новая позиция будет автоматически закреплена за изделием:* `{item_group}`")
        new_col1, new_col2, new_col3 = st.columns(3)
        with new_col1:
            new_name = st.text_input("Название новой позиции:", placeholder="Например: Пошив спец. тесьмы")
        with new_col2:
            new_unit = st.selectbox("Ед. измерения:", ["м.п.", "кв.м.", "шт."])
        with new_col3:
            new_price = st.number_input("Базовая цена (руб):", min_value=0, value=100)
            
        if entry_type == "🧵 Работа":
            new_markup = st.number_input("Базовая наценка (%):", min_value=0, value=0, step=5)
            if st.button("➕ Записать новую работу в Excel", use_container_width=True):
                if new_name.strip():
                    new_row = pd.DataFrame([{"Наименование работы": new_name.strip(), "Ед.изм": new_unit, "Цена (руб)": new_price, "Наценка (%)": new_markup, "Категория": item_group}])
                    pd.concat([df_price, new_row], ignore_index=True).to_excel(PRICE_FILE, index=False)
                    st.rerun()
        else:
            if st.button("➕ Записать новый материал в Excel", use_container_width=True):
                if new_name.strip():
                    new_row = pd.DataFrame([{"Наименование материала": new_name.strip(), "Ед.изм": new_unit, "Цена (руб)": new_price, "Категория": item_group}])
                    pd.concat([df_materials, new_row], ignore_index=True).to_excel(MATERIALS_FILE, index=False)
                    st.rerun()

    elif selected_item:
        unit_type = filtered_dict[selected_item]["Ед.изм"] if selected_item in filtered_dict else "шт."
        base_price = filtered_dict[selected_item]["Цена (руб)"] if selected_item in filtered_dict else 0
        excel_markup = int(filtered_dict[selected_item].get("Наценка (%)", 0)) if selected_item in filtered_dict else 0
        
        calc_col1, calc_col2, calc_col3 = st.columns(3)
        with calc_col1:
            qty = st.number_input(f"Количество ({unit_type}):", min_value=0.01, value=1.0, step=0.1, key="qty_input")
        
        with calc_col2:
            if entry_type == "🧵 Работа":
                markup = st.number_input("Наценка за сложность / бархат (%)", min_value=0, max_value=200, value=excel_markup, step=5)
            else:
                markup = 0
                st.markdown(f"<p style='padding-top:30px; color:gray;'>Цена за единицу: <b>{base_price} руб.</b></p>", unsafe_allow_html=True)
        
        with calc_col3:
            final_unit_price = round(base_price * (1 + markup / 100), 2)
            row_total = round(qty * final_unit_price, 2)
            st.markdown(f"<h4 style='margin:0; padding-top:25px;'>Итог строки: {row_total} руб.</h4>", unsafe_allow_html=True)
            
        if st.button("📥 Добавить эту строчку в смету изделия", use_container_width=True, key="add_row_btn"):
            display_name = f"{selected_item} (+{markup}% нац.)" if markup > 0 else selected_item
            st.session_state.current_items.append({
                "Изделие": item_group,
                "Тип": "Работа" if entry_type == "🧵 Работа" else "Материал",
                "Описание позиции": display_name,
                "Количество": qty,
                "Ед.изм": unit_type,
                "Цена, р": final_unit_price,
                "Стоимость, р": row_total
            })
            st.toast("Добавлено!", icon="✅")
            st.rerun()

st.markdown("---")
st.subheader("📋 Состав и детализация текущего заказа")

if len(st.session_state.current_items) > 0:
    df_items = pd.DataFrame(st.session_state.current_items)
    unique_products = df_items["Изделие"].unique()
    
    for product in unique_products:
        st.markdown(f"#### 📦 Изделие: {product}")
        df_prod = df_items[df_items["Изделие"] == product].drop(columns=["Изделие"])
        st.table(df_prod)
        prod_subtotal = df_prod["Стоимость, р"].sum()
        st.markdown(f"<p style='text-align: right; color: gray;'>Итого по изделию «{product}»: <b>{prod_subtotal:,} р.</b></p>", unsafe_allow_html=True)
    
    st.divider()
    total_sum = df_items["Стоимость, р"].sum()
    st.success(f"### 💰 ВСЕГО К ОПЛАТЕ ПО ВСЕМ ИЗДЕЛИЯМ: {total_sum:,} руб.")
    
    if st.button("🗑️ Очистить всю смету", use_container_width=True, key="clear_all_btn"):
        st.session_state.current_items = []
        st.rerun()
        
    if st.button("💾 Полностью сохранить весь проект в систему", use_container_width=True, key="save_project_btn"):
        st.session_state.orders_db[str(order_digits)] = {
            "project_name": project_name,
            "client_name": client_name,
            "client_phone": client_phone,
            "designer_name": designer_name,
            "admin_name": admin_name,
            "order_date": str(order_date),
            "items": st.session_state.current_items,
            "total_sum": total_sum
        }
        st.balloons()
else:
    st.info("В заказе пока пусто. Выберите изделие наверху, добавьте операции, и здесь сформируется структурированная смета.")
