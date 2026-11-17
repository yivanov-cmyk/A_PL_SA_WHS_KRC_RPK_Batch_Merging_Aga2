#IMPORT MODULES TEMPLATE
import datetime
import logging
import traceback
import os
import openpyxl
import time

#IMPORT UTILS AND OTHER PYTHON FILES
import U_mail
import U_tracking

#IMPORT MODULES FOR PROCESS - WORK FUNCTION (OPTIONAL)
import numpy
from pandas.testing import assert_frame_equal
import pandas
from google.cloud import bigquery
from google.oauth2 import service_account
import warnings
from itertools import combinations

warnings.simplefilter(action='ignore', category=FutureWarning)
pandas.set_option('display.max_columns', None)
pandas.set_option('display.width', None)
pandas.set_option('display.max_colwidth', None)
pandas.set_option('display.float_format', '{:.5f}'.format)

#TEMPLATE VARIABLES - MANUAL
gv_bot_name = "A_PL_SA_WHS_KRC_RPK_Batch_Merging v4.3"
gv_mailing_list = "agandrzejczak@fmlogistic.com,plpythonbot@fmlogistic.com,mklimowicz@fmlogistic.com"
gv_runtype = 1 #1 - PROD #2 - TEST
#TEMPLATE VARIABLES - AUTO
gv_username = os.getlogin()
gv_computer_id = os.environ['COMPUTERNAME']
gv_log_folder = "F_BotLogs\\BL_"+gv_bot_name+"\\"
gv_log_path_for_mail = "F_BotLogs/BL_"+gv_bot_name+"/*"
gv_config_path = "F_Config\\CONFIG_"+gv_bot_name+".xlsx"

#TRIGGERS FOR TEMPLATE
trig_mail_if_fail_1 = True
trig_mail_if_fail_2 = False
trig_mail_if_succesful_1 = False
trig_mail_if_succesful_2 = False
trig_second_run = False
trig_print_debug_info = True

def start():
    #START BLOCK
    global gcp_report_botname
    global gcp_report_status
    global gcp_report_items_total
    global gcp_report_items_succesful
    global gcp_report_items_failed
    global gcp_report_time_start
    global gcp_report_time_end
    global gcp_report_add_info
    gcp_report_botname = gv_bot_name
    gcp_report_status = ""
    gcp_report_items_total = 0
    gcp_report_items_succesful = 0
    gcp_report_items_failed = 0
    gcp_report_time_start = ""
    gcp_report_time_end = ""
    gcp_report_add_info = ""
    if gv_runtype == 1: gcp_report_runtype = "PROD"
    else: gcp_report_runtype = "TEST"
    if not os.path.exists(gv_log_folder):
        os.makedirs(gv_log_folder)
    gv_log_filename = gv_log_folder + "BotLog" + str(datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")) + ".log"
    logging.basicConfig(filename=gv_log_filename, level=logging.INFO)
    print(gv_bot_name + ": start work at: " + str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    print(gv_bot_name + ": logging path: " + gv_log_filename)
    logging.info(gv_bot_name + ": start work at: " + str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    gcp_report_time_start = str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    #WORK BLOCK
    try:
        work()
        gcp_report_status = "WORKED1"
        if trig_mail_if_succesful_1:
            U_mail.start(robot_name=gv_bot_name,mailing_list=gv_mailing_list,log_location=gv_log_path_for_mail,status=gcp_report_status,addinfo=gcp_report_add_info)
    except Exception:
        print(gv_bot_name + ": failed to perform work at its first run")
        traceback.print_exc()
        logging.info(gv_bot_name + ": failed to perform work at its first run")
        error_msg = traceback.format_exc()
        logging.info(error_msg)
        gcp_report_status = "ERROR1"
        if trig_mail_if_fail_1:
            U_mail.start(robot_name=gv_bot_name,mailing_list=gv_mailing_list,log_location=gv_log_path_for_mail,status=gcp_report_status,addinfo=gcp_report_add_info)
        if trig_second_run:
            try:
                work()
                gcp_report_status = "WORKED2"
                if trig_mail_if_succesful_2:
                    gcp_report_add_info = "COMPLETE AT SECOND RUN"
                    U_mail.start(robot_name=gv_bot_name,mailing_list=gv_mailing_list,log_location=gv_log_path_for_mail,status=gcp_report_status,addinfo=gcp_report_add_info)
            except Exception:
                print(gv_bot_name + ": failed to perform work at its second run")
                traceback.print_exc()
                logging.info(gv_bot_name + ": failed to perform work at its second run")
                error_msg = traceback.format_exc()
                logging.info(error_msg)
                gcp_report_status = "ERROR2"
                if trig_mail_if_fail_2:
                    U_mail.start(robot_name=gv_bot_name,mailing_list=gv_mailing_list,log_location=gv_log_path_for_mail,status=gcp_report_status,addinfo=gcp_report_add_info)
    #END BLOCK
    print(gv_bot_name + ": robot complet work at: " + str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    logging.info(gv_bot_name + ": robot complet work at: " + str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    gcp_report_time_end = str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    U_tracking.send_report(v_botname=gcp_report_botname, v_status=gcp_report_status, v_item_total=gcp_report_items_total,
                           v_items_succesful=gcp_report_items_succesful, v_items_failed=gcp_report_items_failed,
                           v_time_start=gcp_report_time_start, v_time_end=gcp_report_time_end,v_user=gv_username,
                           v_computer = gv_computer_id,v_runtype = gcp_report_runtype,v_add_info=gcp_report_add_info)
    print(gv_bot_name + ": send report to GCP")
    logging.info(gv_bot_name + ": send report to GCP")
    for handler in logging.root.handlers[:]: logging.root.removeHandler(handler)
    logging.shutdown()
    logging.root.handlers.clear()


def download_query_history():
    # DOWNLOAD HISTORICAL MOVES IN PICKING FOR CALCULATION
    file = open('F_Configs/SQL_A_PL_SA_WHS_KRC_RPK_Batch_Merging.txt', mode='r')
    query = file.read()
    # print(query)
    file.close()
    credentials = service_account.Credentials.from_service_account_file(
        'F_Tokens/fml-prod-rep-pl-rfx-krc.json', scopes=["https://www.googleapis.com/auth/cloud-platform"])
    clientx = bigquery.Client(credentials=credentials, project=credentials.project_id, )
    dataframe = (clientx.query(query).result().to_dataframe())
    dataframe.to_excel("data_krc_batch_stock.xlsx", index=False)
    #print(dataframe)
    return dataframe

# --- FUNKCJA POMOCNICZA 1 ---
def find_best_single_batch_merge(df_batch, target_fill_rate=1.0, find_best_above_threshold=False):
    """
    Znajduje najlepszą kombinację palet w ramach jednej partii,
    uwzględniając częściowe pobieranie z palet-córek.
    """
    if len(df_batch) < 2:
        return None, 0

    best_combination_df = None
    best_fill_rate = 0
    paletisation = df_batch.iloc[0]['item_qty_lvl30']

    # Pętla po wszystkich możliwych rozmiarach kombinacji (pary, trójki, itd.)
    for r in range(2, len(df_batch) + 1):
        for combo_indices in combinations(df_batch.index, r):
            # Sortujemy kombinację, by największa paleta była "matką"
            combo_df = df_batch.loc[list(combo_indices)].sort_values(by='hd_tot_pcs', ascending=False)

            # --- Inteligentna symulacja z częściowym pobieraniem ---
            mother_pallet = combo_df.iloc[0]
            real_merged_pcs = mother_pallet['hd_tot_pcs']

            # Pętla po "córkach" w tej konkretnej kombinacji
            for _, daughter_pallet in combo_df.iloc[1:].iterrows():
                pcs_needed = paletisation - real_merged_pcs
                if pcs_needed <= 0:
                    break
                pcs_to_take = min(pcs_needed, daughter_pallet['hd_tot_pcs'])
                real_merged_pcs += pcs_to_take

            current_fill_rate = real_merged_pcs / paletisation

            # Oceniamy wynik realistycznej symulacji
            if find_best_above_threshold:
                if current_fill_rate >= 0.7 and current_fill_rate > best_fill_rate:
                    best_fill_rate = current_fill_rate
                    best_combination_df = combo_df
            else:
                if abs(current_fill_rate - target_fill_rate) < 1e-9:
                    return combo_df, target_fill_rate

    return best_combination_df, best_fill_rate


# --- FUNKCJA POMOCNICZA 2 ---
def find_best_two_batch_merge(df_merge, target_fill_rate=1.0, find_best_above_threshold=False):
    """
    Znajduje najlepszą kombinację z dwóch różnych partii, używając pre-konsolidacji.
    """
    if df_merge['ipg_batch'].nunique() < 2:
        return None, 0

    paletisation = df_merge.iloc[0]['item_qty_lvl30']

    # --- Krok 1: Pre-konsolidacja ---
    consolidated_batches = df_merge.groupby('ipg_batch').agg(total_pcs=('hd_tot_pcs', 'sum')).reset_index()

    if len(consolidated_batches) < 2:
        return None, 0

    best_combination_batches = None
    best_fill_rate = 0

    # --- Krok 2: Szukanie najlepszej pary skonsolidowanych grup ---
    for combo_indices in combinations(consolidated_batches.index, 2):
        batch1_info = consolidated_batches.loc[combo_indices[0]]
        batch2_info = consolidated_batches.loc[combo_indices[1]]

        total_pcs = batch1_info['total_pcs'] + batch2_info['total_pcs']

        if total_pcs < 0.7 * paletisation:
            continue

        if batch1_info['total_pcs'] > batch2_info['total_pcs']:
            mother_pcs, daughter_pcs = batch1_info['total_pcs'], batch2_info['total_pcs']
        else:
            mother_pcs, daughter_pcs = batch2_info['total_pcs'], batch1_info['total_pcs']

        pcs_needed = paletisation - mother_pcs
        pcs_to_take = min(pcs_needed, daughter_pcs) if pcs_needed > 0 else 0
        real_merged_pcs = mother_pcs + pcs_to_take
        current_fill_rate = real_merged_pcs / paletisation

        if find_best_above_threshold:
            if current_fill_rate >= 0.7 and current_fill_rate > best_fill_rate:
                best_fill_rate = current_fill_rate
                best_combination_batches = (batch1_info['ipg_batch'], batch2_info['ipg_batch'])
        else:
            if abs(current_fill_rate - 1.0) < 1e-9:
                return (batch1_info['ipg_batch'], batch2_info['ipg_batch']), 1.0

    return best_combination_batches, best_fill_rate


def find_perfect_subset(df_pallets_to_check, target_pcs):
    """
    Przeszukuje wszystkie kombinacje palet w danym DataFrame, aby znaleźć
    podzbiór, którego suma 'hd_tot_pcs' jest dokładnie równa 'target_pcs'.

    Args:
        df_pallets_to_check (DataFrame): Palety z JEDNEJ partii do sprawdzenia.
        target_pcs (int): Dokładna liczba sztuk, której szukamy.

    Returns:
        DataFrame: Podzbiór palet, które sumują się do target_pcs, lub None.
    """
    # Sprawdzamy wszystkie możliwe wielkości kombinacji (od 1 palety do wszystkich)
    for r in range(1, len(df_pallets_to_check) + 1):
        for combo_indices in combinations(df_pallets_to_check.index, r):
            combo_df = df_pallets_to_check.loc[list(combo_indices)]

            # Jeśli suma w kombinacji jest idealna, zwracamy ten podzbiór
            if combo_df['hd_tot_pcs'].sum() == target_pcs:
                return combo_df

    # Jeśli pętla się zakończy i nic nie znajdziemy, zwracamy None
    return None

# --- GŁÓWNA FUNKCJA PRZETWARZAJĄCA ---
# Ta funkcja zawiera całą logikę dla danej kategorii (Standard lub Japan)
def process_pallet_logic(df_cat, params, df_action_list, lv_pair_nb):
    print(f"\n== Rozpoczynam logikę dla kategorii: {params['category_name']} ==")

    # --- Krok 1: Podział na pule robocze ---
    # Definiujemy statusy hold dla różnych pul
    hold_code_split = ' '
    hold_code_merge_mix = 'UNF'

    # Pula MIX: Podzbiór palet MERGE w określonym "średnim" wieku.
    df_for_mix = df_cat[
        (df_cat['l_parameter'] >= params['l_param_mix_min']) & (df_cat['l_parameter'] <= params['l_param_mix_max']) & (
                    df_cat['ipg_grade'] == 4) & (df_cat['ipg_hold_code'] == hold_code_merge_mix)].copy()
    # Pula SPLIT: Palety prawie pełne (>=70%), gotowe do oddania części towaru.
    df_for_split = df_cat[
        (df_cat['hd_percent_of_full_pal'] >= 0.7) &
        (df_cat['l_parameter'] < params['l_param_mix_min']) &
        (df_cat['ipg_grade'] == 4) &
        (df_cat['ipg_hold_code'] == hold_code_split)
        ].copy()
    # Pula MERGE: Palety niepełne (<70%), główni kandydaci do scalania.
    df_for_merge = df_cat[
        (df_cat['hd_percent_of_full_pal'] < 0.7) &
        (df_cat['l_parameter'] < params['l_param_mix_min']) &
        (df_cat['ipg_grade'] == 4) &
        (df_cat['ipg_hold_code'] == hold_code_merge_mix)
        ].copy()

    print(f"  Pule robocze: MIX({len(df_for_mix)}), MERGE({len(df_for_merge)}), SPLIT({len(df_for_split)})")

    # Usuwamy palety MIX z puli MERGE, aby nie były przetwarzane dwa razy.
    df_for_merge = df_for_merge[~df_for_merge['hd_number'].isin(df_for_mix['hd_number'])].copy()

    # --- Krok 2: Nowa, wieloetapowa logika MIX ---
    while not df_for_mix.empty:
        df_for_mix = df_for_mix.sort_values(by=['ipg_batch', 'hd_tot_pcs'], ascending=[True, False]).reset_index(
            drop=True)
        mix_mother = df_for_mix.iloc[0]
        paletisation = mix_mother['item_qty_lvl30']
        print(f"\n  --- Nowa iteracja MIX ---")
        print(
            f"  Matka: {mix_mother['hd_number']}, Zapełnienie: {mix_mother['hd_percent_of_full_pal']:.1%}, Partia: {mix_mother['ipg_batch']}")

        # --- PRÓBA GŁÓWNA: Budowanie grupy MIX + MERGE ---
        current_group = [{'pallet_data': mix_mother.copy(), 'source_df': 'mix'}]
        merged_pcs = mix_mother['hd_tot_pcs']
        used_batches = {mix_mother['ipg_batch']}
        indices_to_drop = {'mix': {0}}

        # Etap 1: Scalanie wewnątrz puli MIX
        if merged_pcs < paletisation:
            for index, daughter_mix in df_for_mix.iloc[1:].iterrows():
                potential_batches = used_batches.union({daughter_mix['ipg_batch']})
                if len(potential_batches) <= 2 and (merged_pcs + daughter_mix['hd_tot_pcs']) <= paletisation:
                    current_group.append({'pallet_data': daughter_mix.copy(), 'source_df': 'mix'})
                    merged_pcs += daughter_mix['hd_tot_pcs']
                    used_batches.add(daughter_mix['ipg_batch'])
                    indices_to_drop.setdefault('mix', set()).add(index)
                    if merged_pcs >= paletisation: break

        # Etap 2: Dopełnienie z puli MERGE
        if merged_pcs < paletisation and len(used_batches) == 1:
            df_for_merge_sorted = df_for_merge.sort_values(by=['ipg_batch', 'hd_tot_pcs'],
                                                           ascending=[True, False])
            for index, daughter_merge in df_for_merge_sorted.iterrows():
                potential_batches = used_batches.union({daughter_merge['ipg_batch']})
                if len(potential_batches) <= 2 and (merged_pcs + daughter_merge['hd_tot_pcs']) <= paletisation:
                    current_group.append({'pallet_data': daughter_merge.copy(), 'source_df': 'merge'})
                    merged_pcs += daughter_merge['hd_tot_pcs']
                    used_batches.add(daughter_merge['ipg_batch'])
                    indices_to_drop.setdefault('merge', set()).add(index)
                    if merged_pcs >= paletisation: break

        final_group = current_group
        final_indices_to_drop = indices_to_drop
        final_merged_pcs = merged_pcs

        # --- Etap 3: Dopełnienie z puli SPLIT (z logiką rezerwową) ---
        if final_merged_pcs / paletisation < 0.7:
            print(f"    -> Grupa MIX+MERGE < 70% ({final_merged_pcs / paletisation:.1%}). Szukam pomocy w SPLIT.")
            df_for_split_sorted = df_for_split.sort_values(by='l_parameter', ascending=False)

            # --- Próba dopełnienia istniejącej grupy ---
            paleta_do_podzialu_1 = None
            for index, split_row in df_for_split_sorted.iterrows():
                if len(used_batches.union({split_row['ipg_batch']})) <= 2:
                    paleta_do_podzialu_1 = split_row
                    break

            if paleta_do_podzialu_1 is not None:
                print(f"    -> Znaleziono SPLIT dla istniejącej grupy. Dopełniam.")
                amount_needed = paletisation - final_merged_pcs
                pcs_to_take = min(amount_needed, paleta_do_podzialu_1['hd_tot_pcs'])
                final_merged_pcs += pcs_to_take
                split_daughter_action = {'pallet_data': paleta_do_podzialu_1.copy(), 'source_df': 'split',
                                         'quantity_to_take': pcs_to_take}
                final_group.append(split_daughter_action)
                original_index = paleta_do_podzialu_1.name
                df_for_split.loc[original_index, 'hd_tot_pcs'] -= pcs_to_take
                #if df_for_split.loc[original_index, 'hd_tot_pcs'] <= 0:
                #    final_indices_to_drop.setdefault('split', set()).add(original_index)
                df_for_split.drop(original_index, inplace=True)
            else:
                # --- STRATEGIA REZERWOWA: Sama Matka + SPLIT ---
                print(f"    -> Brak pasującego SPLIT. Uruchamiam strategię rezerwową.")
                # Resetujemy grupę do samej matki
                group_fallback = [{'pallet_data': mix_mother.copy(), 'source_df': 'mix'}]
                merged_pcs_fallback = mix_mother['hd_tot_pcs']
                used_batches_fallback = {mix_mother['ipg_batch']}
                indices_fallback = {'mix': {0}}

                paleta_do_podzialu_2 = None
                if not df_for_split_sorted.empty:
                    paleta_do_podzialu_2 = df_for_split_sorted.iloc[0]  # Bierzemy dowolną, bo matka ma tylko 1 partię

                if paleta_do_podzialu_2 is not None:
                    amount_needed = paletisation - merged_pcs_fallback
                    pcs_to_take = min(amount_needed, paleta_do_podzialu_2['hd_tot_pcs'])

                    # Sprawdzamy, czy ta strategia rezerwowa w ogóle ma szansę powodzenia
                    if (merged_pcs_fallback + pcs_to_take) / paletisation >= 0.7:
                        print(f"    -> Sukces strategii rezerwowej.")
                        merged_pcs_fallback += pcs_to_take
                        split_daughter_action = {'pallet_data': paleta_do_podzialu_2.copy(), 'source_df': 'split',
                                                 'quantity_to_take': pcs_to_take}
                        group_fallback.append(split_daughter_action)
                        original_index = paleta_do_podzialu_2.name
                        df_for_split.loc[original_index, 'hd_tot_pcs'] -= pcs_to_take
                        #if df_for_split.loc[original_index, 'hd_tot_pcs'] <= 0:
                        #    indices_fallback.setdefault('split', set()).add(original_index)
                        df_for_split.drop(original_index, inplace=True)


                        # Podmieniamy wyniki na te ze strategii rezerwowej
                        final_group = group_fallback
                        final_indices_to_drop = indices_fallback
                        final_merged_pcs = merged_pcs_fallback
                    else:
                        print(f"    -> Strategia rezerwowa nie dała wyniku >= 70%.")
                else:
                    print(f"    -> Strategia rezerwowa nieudana - brak palet SPLIT.")

        # --- FINALNA DECYZJA (na podstawie finalnego wyniku) ---
        if final_merged_pcs / paletisation >= 0.7:
            # SUKCES
            fill_rate = final_merged_pcs / paletisation
            print(
                f"  [MIX SUKCES] Połączono paletę {mix_mother['hd_number']} z {len(final_group) - 1} innymi. Zapełnienie: {fill_rate:.1%}")
            lv_pair_nb += 1
            group_actions = []
            mother_action = final_group[0]['pallet_data'].copy()
            mother_action['action_type'] = f"Łączenie matka >MOTHER_MIX{params['action_suffix']}"
            mother_action['pair_number'] = lv_pair_nb
            mother_action['quantity_to_allocation'] = mother_action['hd_tot_pcs']
            group_actions.append(mother_action)

            temp_merged_pcs = mother_action['hd_tot_pcs']

            for item in final_group[1:]:
                if temp_merged_pcs >= paletisation:
                    break
                daughter_action = item['pallet_data'].copy()
                source = item['source_df']
                daughter_action['action_type'] = f"Łączenie córka >DAUGHTER_{source.upper()}{params['action_suffix']}"
                daughter_action['pair_number'] = lv_pair_nb

                pcs_needed = paletisation - temp_merged_pcs

                if source == 'split':
                    pcs_to_take = min(pcs_needed, item['quantity_to_take'])
                else:
                    pcs_to_take = min(pcs_needed, daughter_action['hd_tot_pcs'])
                daughter_action['quantity_to_allocation'] = pcs_to_take
                temp_merged_pcs += pcs_to_take
                group_actions.append(daughter_action)
            df_action_list = pandas.concat([df_action_list, pandas.DataFrame(group_actions)], ignore_index=True)
            for source, indices in final_indices_to_drop.items():
                if source == 'mix':
                    df_for_mix.drop(list(indices), inplace=True)
                elif source == 'merge':
                    df_for_merge.drop(list(indices), inplace=True)
                elif source == 'split':
                    df_for_split.drop(list(indices), inplace=True)
        else:
            # PORAŻKA
            print(f"  [MIX PORAŻKA] Paleta {mix_mother['hd_number']} nie została połączona po wszystkich próbach.")
            df_for_mix.loc[0, 'action_type'] = f"WAIT OR UTYLIZATION{params['action_suffix']}"
            df_action_list = pandas.concat([df_action_list, df_for_mix.iloc[0:1]], ignore_index=True)
            df_for_mix.drop(0, inplace=True)

    # --- Bramka Kontrolna: Uruchamiaj MERGE/TOP-UP tylko, jeśli jest co najmniej 2 kandydatów ---
    if len(df_for_merge) >= 2:

        # --- Krok 3: Nowa, hierarchiczna logika MERGE ---
        # Pętla działa, dopóki można wykonać jakikolwiek korzystny ruch, zawsze zaczynając od najwyższego priorytetu.
        while True:
            move_made_in_iteration = False

            # Funkcja pomocnicza do zapisu akcji, aktualizacji ilości i usuwania zużytych palet.
            def record_merge_action(combination_df, action_name, mother_rule='largest'):
                nonlocal df_for_merge, df_action_list, lv_pair_nb

                # --- KROK 1: Wybór matki (bez zmian) ---
                if mother_rule == 'largest':
                    mother_pallet_series = combination_df.sort_values(by='hd_tot_pcs', ascending=False).iloc[0]
                elif mother_rule == 'oldest_batch':
                    # Używamy już poprawionej wersji z poprzedniej dyskusji
                    mother_pallet_series = \
                    combination_df.sort_values(by=['ipg_batch', 'hd_tot_pcs'], ascending=[True, False]).iloc[0]

                # --- NOWY KROK 2.1: Oddzielenie matki od córek ---
                mother_index = mother_pallet_series.name
                daughters_df = combination_df.drop(mother_index)

                # --- NOWY KROK 2.2: Optymalne sortowanie córek (od najmniejszej) ---
                # To zapewnia, że zawsze będziemy próbowali zużyć najpierw mniejsze palety.
                sorted_daughters_df = daughters_df.sort_values(by='hd_tot_pcs', ascending=True)

                print(
                    f"  [MERGE SUKCES - {action_name}] Znaleziono kombinację. Matka: {mother_pallet_series['hd_number']}, Córki: {list(sorted_daughters_df['hd_number'])}")

                lv_pair_nb += 1
                group_actions = []
                paletisation = mother_pallet_series['item_qty_lvl30']
                temp_merged_pcs = 0

                # --- KROK 3: Przetworzenie matki (zawsze jako pierwszej) ---
                mother_action = mother_pallet_series.to_dict()
                mother_action['pair_number'] = lv_pair_nb
                mother_action['action_type'] = f"Łączenie matka >{action_name} - MOTHER{params['action_suffix']}"
                mother_action['quantity_to_allocation'] = mother_pallet_series['hd_tot_pcs']
                temp_merged_pcs += mother_pallet_series['hd_tot_pcs']
                group_actions.append(mother_action)

                # --- ZMIENIONY KROK 4: Pętla po OPTYMALNIE POSORTOWANYCH córkach ---
                for daughter_index, daughter_row in sorted_daughters_df.iterrows():
                    # Jeśli paleta jest już pełna, nie ma potrzeby dodawać kolejnych córek
                    if temp_merged_pcs >= paletisation:
                        break

                    action_row = daughter_row.to_dict()
                    action_row['pair_number'] = lv_pair_nb
                    action_row['action_type'] = f"Łączenie córka >{action_name} - DAUGHTER{params['action_suffix']}"

                    pcs_needed = paletisation - temp_merged_pcs
                    pcs_to_take = min(pcs_needed, daughter_row['hd_tot_pcs'])

                    action_row['quantity_to_allocation'] = pcs_to_take
                    temp_merged_pcs += pcs_to_take

                    # Aktualizacja ilości na oryginalnej palecie w df_for_merge
                    df_for_merge.loc[daughter_index, 'hd_tot_pcs'] -= pcs_to_take
                    group_actions.append(action_row)

                df_action_list = pandas.concat([df_action_list, pandas.DataFrame(group_actions)], ignore_index=True)

                # --- KROK 5: Sprzątanie (logika bez zmian) ---
                fully_used_daughters_indices = df_for_merge[
                    df_for_merge.index.isin(daughters_df.index) & (df_for_merge['hd_tot_pcs'] <= 0)
                    ].index

                # Połącz indeks matki i w pełni zużytych córek
                indices_to_drop = list(fully_used_daughters_indices) + [mother_index]
                df_for_merge.drop(indices_to_drop, inplace=True, errors='ignore')

                return True
            # --- Priorytet 1: Szukanie idealnego scalenia 100% z jednej partii ---
            for batch_id in df_for_merge['ipg_batch'].unique():
                df_batch = df_for_merge[df_for_merge['ipg_batch'] == batch_id]
                # Sortujemy palety od najmniejszej do największej, aby faworyzować
                # zużywanie mniejszych palet w całości.
                df_batch = df_batch.sort_values(by='hd_tot_pcs', ascending=True)
                perfect_combo, _ = find_best_single_batch_merge(df_batch, target_fill_rate=1.0)
                if perfect_combo is not None:
                    move_made_in_iteration = record_merge_action(perfect_combo, 'PERFECT MERGER', mother_rule='largest')
                    break
            if move_made_in_iteration: continue

            # --- Priorytet 2: Szukanie idealnego scalenia 100% z dwóch partii ---
            best_move_p2_batches, _ = find_best_two_batch_merge(df_for_merge, target_fill_rate=1.0)
            if best_move_p2_batches is not None:
                # Upewnijmy się, że Partia A jest zawsze starsza
                batch_nums = sorted(list(best_move_p2_batches))
                batch_A_num, batch_B_num = batch_nums[0], batch_nums[1]

                df_batch_A = df_for_merge[df_for_merge['ipg_batch'] == batch_A_num]
                df_batch_B = df_for_merge[df_for_merge['ipg_batch'] == batch_B_num]

                paletisation = df_batch_A.iloc[0]['item_qty_lvl30']

                # Bierzemy całą Partię A jako bazę
                suma_A = df_batch_A['hd_tot_pcs'].sum()
                pcs_needed = paletisation - suma_A

                if pcs_needed > 0:
                    # Szukaj idealnego podzbioru w Partii B
                    perfect_subset_B = find_perfect_subset(df_batch_B, pcs_needed)

                    if perfect_subset_B is not None:
                        # SUKCES! Mamy idealną kombinację
                        print( f"  [PRIORYTET 2 SUKCES] Znaleziono idealne dopasowanie: Cała partia {batch_A_num} + podzbiór z partii {batch_B_num}")
                        final_combination = pandas.concat([df_batch_A, perfect_subset_B])
                    else:
                        print(f"  [PRIORYTET 2 FALLBACK] Brak idealnego podzbioru, przekazuję obie partie do ręcznego scalenia.")
                        final_combination = pandas.concat([df_batch_A, df_batch_B])

                if final_combination is not None:
                    move_made_in_iteration = record_merge_action(final_combination, 'PERFECT MIXED MERGE',
                                                                     'oldest_batch')
                if move_made_in_iteration: continue




            # --- Priorytet 4: Szukanie najlepszego scalenia >=70% z dwóch partii ---
            best_move_p4_batches, _ = find_best_two_batch_merge(df_for_merge, find_best_above_threshold=True)
            if best_move_p4_batches is not None:
                combo_df_p4 = df_for_merge[df_for_merge['ipg_batch'].isin(best_move_p4_batches)]
                move_made_in_iteration = record_merge_action(combo_df_p4, 'MIXED MERGER', mother_rule='oldest_batch')
            if move_made_in_iteration: continue

            # Po zakończeniu pętli `for`, jeśli wykonano jakikolwiek ruch, zrestartuj główny `while`
            if move_made_in_iteration: continue


            # --- Priorytet 3: Szukanie najlepszego scalenia >=70% z jednej partii ---
            batches_to_check = df_for_merge['ipg_batch'].unique()
            for batch_id in batches_to_check:
                # Upewnijmy się, że partia wciąż istnieje w puli (mogła być zużyta w poprzedniej iteracji)
                if batch_id not in df_for_merge['ipg_batch'].unique():
                    continue

                df_batch = df_for_merge[df_for_merge['ipg_batch'] == batch_id]

                # Szukamy najlepszej kombinacji >=70% tylko w tej jednej partii
                best_combo_in_batch, _ = find_best_single_batch_merge(
                    df_batch,
                    find_best_above_threshold=True
                )

                # Jeśli znaleziono poprawną kombinację, wykonaj akcję OD RAZU
                if best_combo_in_batch is not None:
                    print(f"  [PRIORYTET 3 'CHCIWY'] Znaleziono i wykonano połączenie dla partii {batch_id}")
                    # record_merge_action sam usuwa zużyte palety z df_for_merge
                    record_merge_action(best_combo_in_batch, 'REG MERGER', 'largest')
                    # Ustawiamy flagę, by na końcu zrestartować całą główną pętlę while
                    move_made_in_iteration = True

            # Jeśli żaden ruch nie został wykonany, kończymy fazę MERGE.
            if not move_made_in_iteration:
                break

        # --- Krok 4: Faza TOP-UP dla POZOSTAŁYCH PAR ---
        # Ta faza próbuje ratować pary palet, które nie zostały połączone w głównej pętli MERGE.
        df_for_split = df_for_split.sort_values(by='l_parameter', ascending=True)
        while len(df_for_merge) >= 2:
            df_for_merge = df_for_merge.sort_values(by='hd_tot_pcs', ascending=False).reset_index(drop=True)
            group_to_rescue = df_for_merge.iloc[0:2]
            paletisation = group_to_rescue.iloc[0]['item_qty_lvl30']
            group_pcs = group_to_rescue['hd_tot_pcs'].sum()
            used_batches = set(group_to_rescue['ipg_batch'].unique())

            # Sprawdzamy, czy para jest kandydatem do ratunku (ma <70%).
            if group_pcs / paletisation < 0.7:
                paleta_do_podzialu = None
                if not df_for_split.empty:
                    # Szukamy palety SPLIT, która nie naruszy limitu 2 partii.
                    for index, split_row in df_for_split.iterrows():
                        potential_batches = used_batches.union({split_row['ipg_batch']})
                        if len(potential_batches) <= 2:
                            paleta_do_podzialu = split_row.to_frame().T
                            break
                if paleta_do_podzialu is not None:
                    # SUKCES: Znaleziono pomoc, tworzymy akcję TOP-UP.
                    split_pallet_row = paleta_do_podzialu.iloc[0]
                    print(
                        f"  [TOP-UP SUKCES] Grupa palet {list(group_to_rescue['hd_number'])} dopełniona z palety {split_pallet_row['hd_number']}.")
                    lv_pair_nb += 1

                    # --- Przetwarzanie Matki ---
                    mother_pallet = group_to_rescue.iloc[0]
                    mother_action = mother_pallet.to_dict()
                    mother_action['pair_number'] = lv_pair_nb
                    mother_action['action_type'] = f"Łączenie matka >MERGE+TOPUP - MOTHER{params['action_suffix']}"
                    mother_action['quantity_to_allocation'] = mother_pallet['hd_tot_pcs']

                    temp_merged_pcs = mother_pallet['hd_tot_pcs']

                    # --- Przetwarzanie Córki z puli MERGE ---
                    daughter_merge_pallet = group_to_rescue.iloc[1]
                    daughter_action_merge = daughter_merge_pallet.to_dict()
                    daughter_action_merge['pair_number'] = lv_pair_nb
                    daughter_action_merge[
                        'action_type'] = f"Łączenie córka >MERGE+TOPUP - DAUGHTER{params['action_suffix']}"

                    pcs_needed_for_merge = paletisation - temp_merged_pcs
                    pcs_to_take_from_merge = min(pcs_needed_for_merge, daughter_merge_pallet['hd_tot_pcs'])
                    daughter_action_merge['quantity_to_allocation'] = pcs_to_take_from_merge

                    temp_merged_pcs += pcs_to_take_from_merge

                    # --- Przetwarzanie Córki z puli SPLIT ---
                    daughter_action_split = split_pallet_row.to_dict()
                    daughter_action_split['pair_number'] = lv_pair_nb
                    daughter_action_split['action_type'] = f"Łączenie córka >SPLIT FOR TOP-UP{params['action_suffix']}"

                    pcs_needed_for_split = paletisation - temp_merged_pcs
                    pcs_to_take_from_split = min(pcs_needed_for_split, split_pallet_row['hd_tot_pcs'])
                    daughter_action_split['quantity_to_allocation'] = pcs_to_take_from_split

                    df_for_split.loc[split_pallet_row.name, 'hd_tot_pcs'] -= pcs_to_take_from_split

                    # Zapis wszystkich trzech akcji do raportu
                    actions_df = pandas.DataFrame([mother_action, daughter_action_merge, daughter_action_split])
                    df_action_list = pandas.concat([df_action_list, actions_df], ignore_index=True)
                    df_for_merge.drop(group_to_rescue.index, inplace=True)
                else:
                    # PORAŻKA: Nie ma czym dopełnić, para zostaje odrzucona.
                    print(
                        f"  [TOP-UP PORAŻKA] Brak palety w puli SPLIT dla grupy {list(group_to_rescue['hd_number'])}.")
                    rejected_group = group_to_rescue.copy()
                    rejected_group['action_type'] = f"ODRZUCONE MERGE{params['action_suffix']}"
                    df_action_list = pandas.concat([df_action_list, rejected_group], ignore_index=True)
                    df_for_merge.drop(group_to_rescue.index, inplace=True)
            else:
                # Para ma już >=70%, więc nie wymaga TOP-UP. Odrzucamy ją, by nie blokowała puli.
                rejected_group = group_to_rescue.copy()
                rejected_group['action_type'] = f"ODRZUCONE MERGE{params['action_suffix']}"
                df_action_list = pandas.concat([df_action_list, rejected_group], ignore_index=True)
                df_for_merge.drop(group_to_rescue.index, inplace=True)

        # Wyłapujemy ostatnią, pojedynczą paletę, która mogła zostać po przetwarzaniu par.
        if not df_for_merge.empty:
            rejected_pallet = df_for_merge.copy()
            rejected_pallet['action_type'] = f"POMINIĘTA - POJEDYNCZA{params['action_suffix']}"
            df_action_list = pandas.concat([df_action_list, rejected_pallet], ignore_index=True)
    else:
        # Ten blok obsłuży przypadki, gdy w puli MERGE od początku jest 0 lub 1 paleta.
        if not df_for_merge.empty:
            print(
                f"  [INFO] Pula MERGE ma mniej niż 2 palety. Paleta {df_for_merge.iloc[0]['hd_number']} zostanie pominięta.")
            rejected_pallet = df_for_merge.copy()
            rejected_pallet['action_type'] = f"POMINIĘTA - POJEDYNCZA{params['action_suffix']}"
            df_action_list = pandas.concat([df_action_list, rejected_pallet], ignore_index=True)

    return df_action_list, lv_pair_nb


def merging_pallets(dataframe):
    dataframe["hd_tot_pcs"] = dataframe.groupby('hd_number')['ipg_base_lv_qty'].transform('sum')
    dataframe["hd_percent_of_full_pal"] = round(dataframe["hd_tot_pcs"] / dataframe["item_qty_lvl30"], 3)
    dataframe["hd_uniq_ipg"] = dataframe.groupby('hd_number')['ipg_number'].transform('nunique')
    dataframe["hd_uniq_batch"] = dataframe.groupby('hd_number')['ipg_batch'].transform('nunique')
    dataframe["hd_uniq_items"] = dataframe.groupby('hd_number')['item_code'].transform('nunique')
    df_hd_full = dataframe[(dataframe['ipg_onwer'] == 'RCP') & (dataframe['hd_uniq_batch'] == 1) & (
            dataframe['hd_uniq_items'] == 1)].copy()
    df_action_list = pandas.DataFrame(columns=df_hd_full.columns)
    df_action_list["pair_number"] = None
    df_action_list["action_type"] = None
    lv_pair_nb = 0
    processed_pallets = set()

    for dv_item in df_hd_full['item_code'].unique():
        print(f"\n--- Przetwarzanie produktu: {dv_item} ---")
        df_single_item_full = df_hd_full[df_hd_full['item_code'] == dv_item].copy()

        standard_params = {"category_name": "Standard", "l_param_mix_min": 90, "l_param_mix_max": 168,
                           "l_param_oldest": 169, "action_suffix": ""}
        japan_params = {"category_name": "Japan", "l_param_mix_min": 30, "l_param_mix_max": 69, "l_param_oldest": 70,
                        "action_suffix": " (JAPAN)"}

        df_standard = df_single_item_full[df_single_item_full['expiry_group_code'] != 7]
        if not df_standard.empty:
            df_action_list, lv_pair_nb = process_pallet_logic(df_standard, standard_params, df_action_list, lv_pair_nb)

        df_japan = df_single_item_full[df_single_item_full['expiry_group_code'] == 7]
        if not df_japan.empty:
            df_action_list, lv_pair_nb = process_pallet_logic(df_japan, japan_params, df_action_list, lv_pair_nb)

    processed_pallets.update(df_action_list['hd_number'].unique())
    df_remaining = df_hd_full[~df_hd_full['hd_number'].isin(processed_pallets)]
    hold_code_to_check = 'UNF'

    df_oldest_std = df_remaining[
        (df_remaining['expiry_group_code'] != 7) &
        (df_remaining['l_parameter'] >= standard_params['l_param_oldest']) &
        (df_remaining['ipg_hold_code'] == hold_code_to_check) &
        (df_remaining['ipg_grade'] == 4)  # <-- DODANY WARUNEK
        ].copy()

    if not df_oldest_std.empty:
        print(f"[NAJSTARSZE] Znaleziono {len(df_oldest_std)} starych palet Standard do oznaczenia.")
        df_oldest_std['action_type'] = f"DO DECYZJI - UTYLIZACJA{standard_params['action_suffix']}"
        df_action_list = pandas.concat([df_action_list, df_oldest_std], ignore_index=True)

    df_oldest_jpn = df_remaining[
        (df_remaining['expiry_group_code'] == 7) &
        (df_remaining['l_parameter'] >= japan_params['l_param_oldest']) &
        (df_remaining['ipg_hold_code'] == hold_code_to_check) &
        (df_remaining['ipg_grade'] == 4)  # <-- DODANY WARUNEK
        ].copy()

    if not df_oldest_jpn.empty:
        print(f"[NAJSTARSZE] Znaleziono {len(df_oldest_jpn)} starych palet Japan do oznaczenia.")
        df_oldest_jpn['action_type'] = f"DO DECYZJI - UTYLIZACJA{japan_params['action_suffix']}"
        df_action_list = pandas.concat([df_action_list, df_oldest_jpn], ignore_index=True)

    final_columns = ["hd_number", "item_code", "ipg_batch", "ipg_onwer", "ipg_grade", "ipg_hold_code",
                     "ipg_production_date", "expiry_group_code", "hd_loc_full", "hd_tot_pcs", "item_qty_lvl30",
                     "hd_percent_of_full_pal", "l_parameter", "pair_number", "quantity_to_allocation", "action_type"]
    df_action_list = df_action_list.reindex(columns=final_columns)
    df_action_list['Ok/NOK'] = ''
    df_action_list['Komentarze/Rekomendacje'] = ''

    print(df_action_list)
    return df_action_list

def work():
    global gcp_report_items_total
    global gcp_report_items_succesful
    global gcp_report_items_failed
    gcp_report_items_total = 1
    gcp_report_items_succesful = 0
    gcp_report_items_failed = 1

    dataframe = download_query_history() #PROD
    dataframe = pandas.read_excel("data_krc_batch_stock.xlsx", dtype={'hd_number': str}) #TEST
    df_action_list = merging_pallets(dataframe=dataframe)
    df_action_list.to_excel("data_krc_action_list_merging.xlsx", index=False)

    gcp_report_items_total = len(df_action_list)#/2
    gcp_report_items_succesful = len(df_action_list)#/2
    gcp_report_items_failed = 0
    #INFO ABOUT FULLY SUCESSFUL RUN
    print(gv_bot_name + ": ROBOT WORKED PROPERLY")
    logging.info(gv_bot_name + ": ROBOT WORKED PROPERLY")

start()
#work()


