import hashlib
import io
import os
import pickle
import queue
import threading
import time
from datetime import date, timedelta

import pandas as pd
import streamlit as st
from garminconnect import Garmin, GarminConnectAuthenticationError
from openpyxl.chart import LineChart, Reference

# Konfiguracja strony pod urządzenia mobilne i komputery
st.set_page_config(
    page_title="Garmin -> Excel Exporter",
    page_icon="⌚",
    layout="centered",
    initial_sidebar_state="collapsed",
)

SESSION_DIR = os.path.expanduser("~/.garminconnect_streamlit_sessions")

ACTIVITY_MAP = {
    "Wszystkie aktywności": "all",
    "Bieganie (wszystkie rodzaje)": "running",
    "Spacer": "walking",
    "Jazda na rowerze": "cycling",
    "Trening siłowy": "strength_training",
    "Wędrówka / Trekking": "hiking",
    "Pływanie": "swimming",
    "Kardio / Sprzęt fitness": "fitness_equipment",
}


# --- FUNKCJE POMOCNICZE I SESJA ---
def get_session_filepath(email: str, password: str) -> str:
    """Generuje unikalną ścieżkę sesji powiązaną z konkretnym loginem i hasłem."""
    os.makedirs(SESSION_DIR, exist_ok=True)
    cred_hash = hashlib.sha256(f"{email.strip().lower()}:{password}".encode("utf-8")).hexdigest()[:24]
    return os.path.join(SESSION_DIR, f"session_{cred_hash}.pkl")


def load_saved_session(email: str, password: str):
    filepath = get_session_filepath(email, password)
    if os.path.exists(filepath):
        try:
            with open(filepath, "rb") as f:
                client = pickle.load(f)
            # Szybki test czy sesja nie wygasła
            client.get_full_name()
            return client
        except Exception:
            pass
    return None


def save_session(client, email: str, password: str):
    try:
        filepath = get_session_filepath(email, password)
        client.prompt_mfa = None
        with open(filepath, "wb") as f:
            pickle.dump(client, f)
    except Exception:
        pass


def start_login_thread(email: str, password: str):
    """Uruchamia logowanie w tle, pozwalając na wpisanie kodu 2FA z maila w interfejsie Streamlit."""
    auth_state = {
        "mfa_needed": threading.Event(),
        "login_done": threading.Event(),
        "mfa_queue": queue.Queue(),
        "client": None,
        "error": None,
        "email": email,
        "password": password,
    }

    def prompt_mfa_callback():
        auth_state["mfa_needed"].set()
        # Czeka do 5 minut na wpisanie kodu przez użytkownika na telefonie
        code = auth_state["mfa_queue"].get(timeout=300)
        return code.strip()

    def worker():
        try:
            client = Garmin(email=email, password=password, prompt_mfa=prompt_mfa_callback)
            client.login()
            auth_state["client"] = client
        except Exception as e:
            auth_state["error"] = e
        finally:
            auth_state["login_done"].set()

    thread = threading.Thread(target=worker, daemon=True)
    auth_state["thread"] = thread
    st.session_state["auth_worker"] = auth_state
    thread.start()


# --- PRZETWARZANIE DANYCH GARMIN ---
def matches_activity_type(activity, selected_key):
    if selected_key == "all":
        return True
    act_type_dto = activity.get("activityType", {})
    type_key = (act_type_dto.get("typeKey") or "").lower()

    if selected_key == "running":
        return "running" in type_key or type_key == "treadmill_running"
    if selected_key == "cycling":
        return "cycling" in type_key or "biking" in type_key
    if selected_key == "walking":
        return "walking" in type_key
    return type_key == selected_key


def format_pace(speed_m_s):
    if not speed_m_s or speed_m_s <= 0:
        return "-"
    pace_seconds = 1000.0 / speed_m_s
    minutes = int(pace_seconds // 60)
    seconds = int(round(pace_seconds % 60))
    if seconds == 60:
        minutes += 1
        seconds = 0
    return f"{minutes}:{seconds:02d}"


def pace_decimal(speed_m_s):
    if not speed_m_s or speed_m_s <= 0:
        return None
    pace_min = (1000.0 / speed_m_s) / 60.0
    return round(min(pace_min, 20.0), 2)


def safe_round(val, digits=2):
    if val is None:
        return None
    try:
        return round(float(val), digits)
    except (ValueError, TypeError):
        return None


def extract_summary_row(act):
    dist_km = (act.get("distance") or 0) / 1000.0
    dur_sec = act.get("duration") or 0
    moving_sec = act.get("movingDuration") or 0
    avg_speed = act.get("averageSpeed") or 0
    max_speed = act.get("maxSpeed") or 0
    gap_speed = act.get("avgGradeAdjustedSpeed") or 0

    return {
        "ID Aktywności": act.get("activityId"),
        "Data i godzina": act.get("startTimeLocal"),
        "Nazwa aktywności": act.get("activityName"),
        "Typ aktywności": act.get("activityType", {}).get("typeKey"),
        "Dystans [km]": safe_round(dist_km, 2),
        "Czas całkowity [min]": safe_round(dur_sec / 60.0, 2),
        "Czas w ruchu [min]": safe_round(moving_sec / 60.0, 2),
        "Czas (HH:MM:SS)": str(timedelta(seconds=int(dur_sec))),
        "Średnie tempo [min/km]": format_pace(avg_speed),
        "Maks. tempo [min/km]": format_pace(max_speed),
        "Tempo GAP (skorygowane o nachylenie)": format_pace(gap_speed),
        "Średnia prędkość [km/h]": safe_round(avg_speed * 3.6, 2),
        "Maks. prędkość [km/h]": safe_round(max_speed * 3.6, 2),
        "Średnie tętno [bpm]": act.get("averageHR"),
        "Maks. tętno [bpm]": act.get("maxHR"),
        "Średnia kadencja [krok/min]": safe_round(act.get("averageRunningCadenceInStepsPerMinute"), 1),
        "Maks. kadencja [krok/min]": safe_round(act.get("maxRunningCadenceInStepsPerMinute"), 1),
        "Liczba kroków": act.get("steps"),
        "Śr. długość kroku [cm]": safe_round(act.get("avgStrideLength"), 1),
        "Śr. oscylacja pionowa [cm]": safe_round(act.get("avgVerticalOscillation"), 2),
        "Śr. odchylenie do długości [%]": safe_round(act.get("avgVerticalRatio"), 2),
        "Śr. czas kontaktu z podłożem GCT [ms]": safe_round(act.get("avgGroundContactTime"), 1),
        "Średnia moc [W]": safe_round(act.get("avgPower"), 1),
        "Maks. moc [W]": safe_round(act.get("maxPower"), 1),
        "Moc znormalizowana NP [W]": safe_round(act.get("normPower"), 1),
        "VO2 Max": act.get("vO2MaxValue"),
        "Spalone kalorie [kcal]": act.get("calories"),
        "Szacowany ubytek wody [ml]": act.get("waterEstimated"),
        "Śr. częstość oddechu [oddech/min]": safe_round(act.get("avgRespirationRate"), 1),
        "Maks. częstość oddechu [oddech/min]": safe_round(act.get("maxRespirationRate"), 1),
        "Wznios (Przewyższenie) [m]": safe_round(act.get("elevationGain"), 1),
        "Spadek terenu [m]": safe_round(act.get("elevationLoss"), 1),
        "Min. wysokość [m]": safe_round(act.get("minElevation"), 1),
        "Maks. wysokość [m]": safe_round(act.get("maxElevation"), 1),
        "Efekt treningu (Główny cel)": act.get("trainingEffectLabel"),
        "Efekt treningu (Aerobowy)": safe_round(act.get("aerobicTrainingEffect"), 1),
        "Efekt treningu (Anaerobowy)": safe_round(act.get("anaerobicTrainingEffect"), 1),
        "Obciążenie treningowe (Load)": safe_round(act.get("activityTrainingLoad"), 1),
        "Strefa tętna 1 [min]": safe_round((act.get("hrTimeInZone_1") or 0) / 60.0, 2),
        "Strefa tętna 2 [min]": safe_round((act.get("hrTimeInZone_2") or 0) / 60.0, 2),
        "Strefa tętna 3 [min]": safe_round((act.get("hrTimeInZone_3") or 0) / 60.0, 2),
        "Strefa tętna 4 [min]": safe_round((act.get("hrTimeInZone_4") or 0) / 60.0, 2),
        "Strefa tętna 5 [min]": safe_round((act.get("hrTimeInZone_5") or 0) / 60.0, 2),
    }


def extract_laps(client, act_id, act_date, act_name):
    lap_rows = []
    try:
        splits_data = client.get_activity_splits(act_id)
        laps = splits_data.get("lapDTOs", [])
        for idx, lap in enumerate(laps, start=1):
            dist_km = (lap.get("distance") or 0) / 1000.0
            dur_sec = lap.get("duration") or 0
            avg_speed = lap.get("averageSpeed") or 0
            max_speed = lap.get("maxSpeed") or 0

            lap_rows.append({
                "ID Aktywności": act_id,
                "Data i godzina": act_date,
                "Nazwa aktywności": act_name,
                "Nr okrążenia": lap.get("lapIndex") or idx,
                "Typ odcinka": lap.get("intensityType"),
                "Dystans [km]": safe_round(dist_km, 3),
                "Czas [s]": safe_round(dur_sec, 1),
                "Czas (MM:SS)": str(timedelta(seconds=int(dur_sec))),
                "Średnie tempo [min/km]": format_pace(avg_speed),
                "Maks. tempo [min/km]": format_pace(max_speed),
                "Śr. prędkość [km/h]": safe_round(avg_speed * 3.6, 2),
                "Średnie tętno [bpm]": lap.get("averageHR"),
                "Maks. tętno [bpm]": lap.get("maxHR"),
                "Średnia kadencja [krok/min]": safe_round(
                    lap.get("averageRunCadence") or lap.get("averageRunningCadenceInStepsPerMinute"), 1
                ),
                "Maks. kadencja [krok/min]": safe_round(lap.get("maxRunCadence"), 1),
                "Długość kroku [cm]": safe_round(lap.get("strideLength"), 1),
                "Czas kontaktu z podłożem GCT [ms]": safe_round(lap.get("groundContactTime"), 1),
                "Oscylacja pionowa [cm]": safe_round(lap.get("verticalOscillation"), 2),
                "Odchylenie pionowe [%]": safe_round(lap.get("verticalRatio"), 2),
                "Średnia moc [W]": safe_round(lap.get("averagePower"), 1),
                "Maks. moc [W]": safe_round(lap.get("maxPower"), 1),
                "Wznios [m]": safe_round(lap.get("elevationGain"), 1),
                "Spadek [m]": safe_round(lap.get("elevationLoss"), 1),
                "Spalone kalorie [kcal]": lap.get("calories"),
            })
    except Exception:
        pass
    return lap_rows


def extract_chart_timeseries(client, act_id, act_date, act_name):
    ts_rows = []
    try:
        details = client.get_activity_details(act_id, maxchart=1000)
        descriptors = details.get("metricDescriptors", [])
        metrics_list = details.get("activityDetailMetrics", [])
        key_to_idx = {d.get("key"): d.get("metricsIndex") for d in descriptors if "key" in d}

        def get_val(arr, key):
            idx = key_to_idx.get(key)
            if idx is not None and idx < len(arr):
                return arr[idx]
            return None

        for item in metrics_list:
            m = item.get("metrics", [])
            dur_sec = get_val(m, "sumElapsedDuration")
            if dur_sec is None:
                dur_sec = get_val(m, "sumDuration") or 0

            dist_m = get_val(m, "sumDistance") or 0
            speed_ms = get_val(m, "directSpeed")
            hr = get_val(m, "directHeartRate")
            cad = get_val(m, "directDoubleCadence")
            if cad is None:
                single_cad = get_val(m, "directRunCadence")
                cad = single_cad * 2 if single_cad else None

            elev = get_val(m, "directElevation")
            power = get_val(m, "directPower")
            gct = get_val(m, "directGroundContactTime")
            vo = get_val(m, "directVerticalOscillation")
            stride = get_val(m, "directStrideLength")

            ts_rows.append({
                "ID Aktywności": act_id,
                "Data": act_date,
                "Nazwa aktywności": act_name,
                "Dystans [km]": safe_round(dist_m / 1000.0, 3),
                "Czas [min]": safe_round(dur_sec / 60.0, 2),
                "Czas (HH:MM:SS)": str(timedelta(seconds=int(dur_sec))),
                "Tętno [bpm]": safe_round(hr, 0),
                "Prędkość [km/h]": safe_round(speed_ms * 3.6, 2) if speed_ms is not None else None,
                "Tempo [min/km]": pace_decimal(speed_ms),
                "Tempo (MM:SS)": format_pace(speed_ms),
                "Kadencja [krok/min]": safe_round(cad, 0),
                "Wysokość [m]": safe_round(elev, 1),
                "Moc [W]": safe_round(power, 0),
                "Kontakt z podłożem GCT [ms]": safe_round(gct, 0),
                "Oscylacja pionowa [cm]": safe_round(vo, 2),
                "Długość kroku [cm]": safe_round(stride, 1),
            })
    except Exception:
        pass
    return ts_rows


def build_excel_charts(workbook, activity_ranges):
    ws_details = workbook["Dane_Wykresow"]
    ws_charts = workbook.create_sheet(title="Wykresy")

    current_row = 2
    for info in activity_ranges:
        r_start = info["start_row"]
        r_end = info["end_row"]
        title_prefix = f"{info['date']} | {info['name']}"

        if r_end - r_start < 3:
            continue

        cats_ref = Reference(ws_details, min_col=4, min_row=r_start, max_row=r_end)

        # 1. Wykres Tętna
        chart_hr = LineChart()
        chart_hr.title = f"{title_prefix} - Tętno [bpm]"
        chart_hr.style = 10
        chart_hr.y_axis.title = "Tętno [bpm]"
        chart_hr.x_axis.title = "Dystans [km]"
        chart_hr.width = 16
        chart_hr.height = 10
        chart_hr.legend = None
        hr_ref = Reference(ws_details, min_col=7, min_row=r_start, max_row=r_end)
        chart_hr.add_data(hr_ref, titles_from_data=False)
        chart_hr.set_categories(cats_ref)
        ws_charts.add_chart(chart_hr, f"B{current_row}")

        # 2. Wykres Prędkości i Kadencji
        chart_spd = LineChart()
        chart_spd.title = f"{title_prefix} - Prędkość [km/h] i Kadencja [spm]"
        chart_spd.style = 12
        chart_spd.y_axis.title = "Wartość"
        chart_spd.x_axis.title = "Dystans [km]"
        chart_spd.width = 16
        chart_spd.height = 10
        spd_ref = Reference(ws_details, min_col=8, min_row=r_start, max_row=r_end)
        cad_ref = Reference(ws_details, min_col=11, min_row=r_start, max_row=r_end)
        chart_spd.add_data(spd_ref, titles_from_data=False)
        chart_spd.add_data(cad_ref, titles_from_data=False)
        chart_spd.set_categories(cats_ref)
        ws_charts.add_chart(chart_spd, f"L{current_row}")

        # 3. Wykres Wysokości i Mocy
        chart_elev = LineChart()
        chart_elev.title = f"{title_prefix} - Wysokość [m] i Moc [W]"
        chart_elev.style = 13
        chart_elev.y_axis.title = "Wartość"
        chart_elev.x_axis.title = "Dystans [km]"
        chart_elev.width = 16
        chart_elev.height = 10
        elev_ref = Reference(ws_details, min_col=12, min_row=r_start, max_row=r_end)
        pwr_ref = Reference(ws_details, min_col=13, min_row=r_start, max_row=r_end)
        chart_elev.add_data(elev_ref, titles_from_data=False)
        chart_elev.add_data(pwr_ref, titles_from_data=False)
        chart_elev.set_categories(cats_ref)
        ws_charts.add_chart(chart_elev, f"V{current_row}")

        current_row += 21


# --- INTERFEJS UŻYTKOWNIKA STREAMLIT ---
st.title("⌚ Garmin Connect -> Excel")
st.caption("Pobieraj pełne statystyki, okrążenia i wykresy bezpośrednio na telefon lub komputer.")

if "garmin_client" not in st.session_state:
    st.session_state["garmin_client"] = None
if "logged_email" not in st.session_state:
    st.session_state["logged_email"] = None
if "awaiting_mfa" not in st.session_state:
    st.session_state["awaiting_mfa"] = False

# KROK 1: LOGOWANIE
if st.session_state["garmin_client"] is None:
    st.subheader("1. Logowanie do Garmin Connect")

    if not st.session_state["awaiting_mfa"]:
        with st.form("login_form"):
            email_in = st.text_input("E-mail Garmin Connect")
            password_in = st.text_input("Hasło", type="password")
            remember_in = st.checkbox("Zapamiętaj moją sesję na tym urządzeniu", value=True)
            submit_login = st.form_submit_button("🔐 Zaloguj się", use_container_width=True)

        if submit_login:
            if not email_in or not password_in:
                st.error("Wpisz adres e-mail oraz hasło.")
            else:
                with st.spinner("Łączenie z Garmin Connect..."):
                    saved_client = load_saved_session(email_in, password_in) if remember_in else None
                    if saved_client is not None:
                        st.session_state["garmin_client"] = saved_client
                        st.session_state["logged_email"] = email_in
                        st.rerun()
                    else:
                        st.session_state["remember_pref"] = remember_in
                        start_login_thread(email_in, password_in)
                        worker = st.session_state["auth_worker"]

                        # Czekaj aż wątek zakończy logowanie lub poprosi o kod 2FA
                        while not worker["login_done"].is_set() and not worker["mfa_needed"].is_set():
                            time.sleep(0.2)

                        if worker["mfa_needed"].is_set() and not worker["login_done"].is_set():
                            st.session_state["awaiting_mfa"] = True
                            st.rerun()
                        elif worker["error"] is not None:
                            st.error(f"Błąd logowania: {worker['error']}")
                        elif worker["client"] is not None:
                            if remember_in:
                                save_session(worker["client"], email_in, password_in)
                            st.session_state["garmin_client"] = worker["client"]
                            st.session_state["logged_email"] = email_in
                            st.rerun()

    else:
        st.info("📩 Garmin wysłał kod weryfikacyjny (2FA) na Twój e-mail lub SMS. Wpisz go poniżej:")
        with st.form("mfa_form"):
            mfa_code = st.text_input("Kod weryfikacyjny (6 cyfr)")
            col1, col2 = st.columns(2)
            with col1:
                submit_mfa = st.form_submit_button("✅ Potwierdź kod", use_container_width=True)
            with col2:
                cancel_mfa = st.form_submit_button("Anuluj", use_container_width=True)

        if cancel_mfa:
            st.session_state["awaiting_mfa"] = False
            st.session_state.pop("auth_worker", None)
            st.rerun()

        if submit_mfa:
            worker = st.session_state.get("auth_worker")
            if not worker:
                st.session_state["awaiting_mfa"] = False
                st.rerun()
            else:
                with st.spinner("Weryfikacja kodu..."):
                    worker["mfa_queue"].put(mfa_code)
                    worker["login_done"].wait(timeout=30)
                    st.session_state["awaiting_mfa"] = False

                    if worker["error"] is not None:
                        st.error(f"Błąd weryfikacji: {worker['error']}")
                    elif worker["client"] is not None:
                        if st.session_state.get("remember_pref", True):
                            save_session(worker["client"], worker["email"], worker["password"])
                        st.session_state["garmin_client"] = worker["client"]
                        st.session_state["logged_email"] = worker["email"]
                        st.rerun()

# KROK 2: POBIERANIE DANYCH (GDY ZALOGOWANO)
else:
    col_info, col_logout = st.columns([3, 1])
    with col_info:
        st.success(f"Zalogowano: **{st.session_state['logged_email']}**")
    with col_logout:
        if st.button("Wyloguj", use_container_width=True):
            st.session_state["garmin_client"] = None
            st.session_state["logged_email"] = None
            st.session_state.pop("excel_bytes", None)
            st.rerun()

    st.subheader("2. Zakres dat i parametry eksportu")

    col_d1, col_d2 = st.columns(2)
    with col_d1:
        start_dt = st.date_input("Data początkowa", value=date.today() - timedelta(days=30))
    with col_d2:
        end_dt = st.date_input("Data końcowa", value=date.today())

    selected_label = st.selectbox("Typ aktywności", options=list(ACTIVITY_MAP.keys()), index=1)
    selected_key = ACTIVITY_MAP[selected_label]

    include_laps = st.checkbox("Pobierz szczegółowe okrążenia / interwały (Laps)", value=True)
    include_charts = st.checkbox("Pobierz przebieg treningu i utwórz wykresy w Excelu", value=True)

    if st.button("🚀 Pobierz dane i przygotuj plik Excel", type="primary", use_container_width=True):
        if start_dt > end_dt:
            st.warning("Data początkowa nie może być późniejsza niż końcowa.")
        else:
            client = st.session_state["garmin_client"]
            status_box = st.empty()
            progress_bar = st.progress(0)

            try:
                status_box.info("Pobieranie listy aktywności z Garmin Connect...")
                raw_activities = client.get_activities_by_date(
                    start_dt.strftime("%Y-%m-%d"), end_dt.strftime("%Y-%m-%d")
                )

                filtered = [
                    act for act in raw_activities if matches_activity_type(act, selected_key)
                ]

                if not filtered:
                    progress_bar.empty()
                    status_box.warning("Nie znaleziono żadnych treningów w wybranym przedziale i kategorii.")
                else:
                    total = len(filtered)
                    summary_rows = []
                    all_laps = []
                    all_timeseries = []
                    activity_ranges = []
                    current_ts_row = 2

                    for idx, act in enumerate(filtered, start=1):
                        act_id = act.get("activityId")
                        act_date = (act.get("startTimeLocal") or "")[:16]
                        act_name = act.get("activityName") or "Trening"

                        status_box.info(f"Pobieranie treningu {idx}/{total}: **{act_name}** ({act_date})...")
                        summary_rows.append(extract_summary_row(act))

                        if include_laps and act_id:
                            all_laps.extend(extract_laps(client, act_id, act_date, act_name))

                        if include_charts and act_id:
                            ts_points = extract_chart_timeseries(client, act_id, act_date, act_name)
                            if ts_points:
                                start_r = current_ts_row
                                end_r = current_ts_row + len(ts_points) - 1
                                activity_ranges.append({
                                    "id": act_id,
                                    "date": act_date,
                                    "name": act_name,
                                    "start_row": start_r,
                                    "end_row": end_r,
                                })
                                all_timeseries.extend(ts_points)
                                current_ts_row = end_r + 1

                        progress_bar.progress(idx / total)

                    status_box.info("Generowanie pliku Excel i rysowanie wykresów...")

                    df_summary = pd.DataFrame(summary_rows)
                    df_raw_all = pd.json_normalize(filtered)

                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        df_summary.to_excel(writer, sheet_name="Podsumowanie", index=False)
                        if all_laps:
                            pd.DataFrame(all_laps).to_excel(writer, sheet_name="Okrążenia", index=False)
                        if all_timeseries:
                            pd.DataFrame(all_timeseries).to_excel(writer, sheet_name="Dane_Wykresow", index=False)
                            try:
                                build_excel_charts(writer.book, activity_ranges)
                            except Exception:
                                pass
                        df_raw_all.to_excel(writer, sheet_name="Wszystkie_Pola_API", index=False)

                    st.session_state["excel_bytes"] = output.getvalue()
                    st.session_state["excel_filename"] = f"garmin_{selected_key}_{start_dt}_{end_dt}.xlsx"
                    st.session_state["df_summary_preview"] = df_summary
                    st.session_state["df_laps_preview"] = pd.DataFrame(all_laps) if all_laps else None

                    status_box.success(f"Gotowe! Pobrano {total} treningów.")

            except GarminConnectAuthenticationError:
                st.error("Sesja wygasła lub wystąpił błąd autoryzacji. Wyloguj się i zaloguj ponownie.")
            except Exception as e:
                st.error(f"Wystąpił błąd: {e}")

    # KROK 3: PRZYCISK POBIERANIA PLIKU EXCEL I PODGLĄD NA TELEFONIE
    if "excel_bytes" in st.session_state:
        st.divider()
        st.download_button(
            label="📥 Pobierz gotowy plik Excel (.xlsx)",
            data=st.session_state["excel_bytes"],
            file_name=st.session_state["excel_filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        with st.expander("📊 Szybki podgląd tabel w przeglądarce"):
            tab1, tab2 = st.tabs(["Podsumowanie", "Okrążenia"])
            with tab1:
                st.dataframe(st.session_state["df_summary_preview"], use_container_width=True)
            with tab2:
                if st.session_state.get("df_laps_preview") is not None:
                    st.dataframe(st.session_state["df_laps_preview"], use_container_width=True)
                else:
                    st.write("Brak pobranych okrążeń.")