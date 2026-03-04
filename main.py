from fastapi import FastAPI, HTTPException
from docx import Document
import httpx
import os
from datetime import datetime

app = FastAPI()

# === Bitrix24 настройки ===
WEBHOOK = "https://izyskaniya.bitrix24.ru/rest/1005/1eahl4oku941fawb/"
DISK_FOLDER_ID = "1706930"
KP_LINK_FIELD = "UF_CRM_1761672766812"  # Поле для ссылки на КП

def format_cost(cost_str):
    try:
        return f"{int(float(cost_str)):,}".replace(",", " ")
    except (ValueError, TypeError):
        return "0"

def build_igi_section(data):
    lines = [
        "Комплекс инженерно-геологических изысканий:",
        f"1.1. Полевые работы:",
        f"— бурение скважин глубиной по {data['igi_drilling_depth']} м – {data['igi_boreholes']} скважин;",
        f"— статическое зондирование грунтов – {data['igi_sounding_points']} точки;",
        f"— гидрогеологические наблюдения при бурении скважин;",
        f"— отбор проб грунтов для лабораторных испытаний;",
        f"— планово-высотная привязка скважин.",
        f"1.2. Лабораторные исследования:",
        f"— определение физических характеристик грунтов;",
        f"— определение механических характеристик грунтов;",
        f"— определение коррозийной активности грунтов;",
        f"— определение химического состава подземных вод.",
        f"1.3. Камеральные работы:",
        f"— статистическая обработка данных лабораторных исследований грунтов;",
        f"— обработка данных полевых испытаний грунтов;",
        f"— разработка технического отчета по инженерно-геологическим изысканиям.",
        f"Сроки выполнения работ: {data['igi_duration_days']} рабочих дней.",
        f"Стоимость работ по п.1 составит: {data['igi_cost']} руб. с учетом НДС."
    ]
    return "\n".join(lines)

def build_igdi_section(data):
    lines = [
        "Комплекс инженерно-геодезических изысканий:",
        f"2.1. Топографо-геодезическая съемка участка, общей площадью {data['igdi_area_ha']} Га;",
        f"2.2. Выполнение в пределах границ топографической съемки полевого обследования и съемки всех наземных, надземных и подземных коммуникаций, определение их типа, характеристики, глубины заложения;",
        f"2.3. Камеральная обработка результатов съемки и оформление топографического плана в масштабе {data['igdi_scale']}, сечение рельефа через {data['igdi_contour_interval']} м;",
        f"2.4. Согласование полноты и достоверности нанесения на топографические планы коммуникаций в эксплуатирующих организациях.",
        f"2.5. Составление технического отчета по результатам выполненных инженерно-геодезических изысканий.",
        f"Сроки выполнения работ: {data['igdi_duration_days']} рабочих дней (в том числе:",
        f"— топографическая съёмка и оформление топографического плана – {data['igdi_survey_days']} рабочих дней;",
        f"— согласование топографического плана с балансодержателями и составление технического отчета – {data['igdi_coordination_days']} рабочих дней).",
        f"Стоимость работ по п. 2: {data['igdi_cost']} руб. с учетом НДС (в том числе:",
        f"— топографическая съёмка и оформление топографического плана – {data['igdi_survey_cost']} руб.;",
        f"— согласование топографического плана с балансодержателями – {data['igdi_coordination_cost']} руб.;",
        f"— составление технического отчета – {data['igdi_report_cost']} руб.).",
        "* оплата счетов балансодержателей не включена в стоимость процедуры согласования. Счета от балансодержателей оплачивает Заказчик.",
        "* В составе работ подача в ИСОГД"
    ]
    return "\n".join(lines)

def build_iei_section(data):
    lines = [
        "Комплекс инженерно-экологических изысканий:",
        f"3.1. Полевые работы:",
        f"- маршрутные наблюдения {data['iei_area_ha']} га;",
        f"- пешеходная гамма-съемка участка {data['iei_area_ha']} га;",
        f"- замеры МЭД гамма-излучения - суммарно {data['iei_gamma_points']} точек в границах участка изысканий;",
        f"- исследования влияния физических факторов (шум - {data['iei_noise_points']} точки, ЭМИ - {data['iei_emi_points']} точки);",
        f"- отбор {data['iei_soil_samples']} объединенных проб почв с глубины 0,0-0,2 м для проведения санитарно-химического (определение содержания тяжелых металлов, бенз(а)пирена и нефтепродуктов) исследования;",
        f"- отбор {data['iei_bio_samples']} объединенных проб почв с глубины 0,0-0,2 м для проведения бактериологического (определение БГКП, энтерококков, патогенных бактерий, в т.ч. сальмонелл) и паразитологического (определение яиц и личинок гельминтов, цист патогенных простейших) и энтомологического (куколки и личинки синантропных мух) исследования;",
        f"- послойный отбор {data['iei_layered_samples_deep']} проб грунта из {data['iei_deep_boreholes']} геолог. скважин (послойно с глубин 0,2-1,0 м, 1,0-2,0 м, ..., 14,0-15,0 м) для проведения санитарно-химического (определение содержания тяжелых металлов, бенз(а)пирена и нефтепродуктов) анализа;",
        f"- послойный отбор {data['iei_layered_samples_shallow']} пробы грунта из {data['iei_shallow_boreholes']} геолог. скважин (послойно с глубин 0,2-1,0 м) для проведения санитарно-химического (определение содержания тяжелых металлов, бенз(а)пирена и нефтепродуктов) анализа;",
        f"- отбор {data['iei_background_soil_samples']} проб почв с глубины 0,0-0,2 м, 0,2-1,0 м, ..., 14,0-15,0 м для проведения санитарно-химического (определение нефтепродуктов фоновая проба) исследования;",
        f"- послойный отбор {data['iei_agro_samples']} проб грунта из {data['iei_pits']} шурфа (послойно с глубин 0,0-0,2; 0,2-0,4; 0,4-0,6; 0,6-0,8; 0,8-1,0 м) для проведения агрохимического анализа;",
        f"- отбор {data['iei_rad_samples']} объединенных проб почв с глубины 0,0-0,2 м для проведения радиационного исследования;",
        f"- отбор {data['iei_water_samples']} пробы подземной воды (при вскрытии) из {data['iei_water_boreholes']} скважины;",
        f"- отбор {data['iei_surface_water_samples']} пробы воды из поверхностного водного объекта;",
        f"- отбор {data['iei_sediment_samples']} пробы донных отложений из водного объекта;",
        f"3.2. Лабораторные исследования проб почв, грунтов;",
        f"3.3. Камеральная обработка данных и разработка технического отчета об инженерно-экологических изысканиях;",
        f"3.4. Получение справки Росгидромета о фоновых концентрациях загрязняющих веществ в атмосферном воздухе (диоксид азота, оксид углерода, серы диоксид и взвешенные вещества);",
        f"3.5. Получение справки Росгидромета о краткой климатической характеристике (для расчета рассеивания ЗВ в атмосферном воздухе);",
        f"3.6. Получение справок о наличии/отсутствии:",
        "- особо охраняемых природных территорий;",
        "- объектов культурного наследия;",
        "- скотомогильников и их санитарных зон;",
        "- краснокнижных видов животных и растений;",
        "- водоохранных и прибрежно-защитных полос водных объектов;",
        "- зон санитарной охраны источников питьевого и хозяйственно-бытового водоснабжения;",
        "- защитных лесов;",
        "- курортных и рекреационных зон;",
        "- о попадании (не попадании) объекта в границы СЗЗ других объектов;",
        "- полигонов ТКО и свалок;",
        "- месторождений полезных ископаемых.",
        f"",
        f"Сроки выполнения работ: {data['iei_duration_days']} рабочих дней.",
        f"Стоимость работ по п.3 составит: {data['iei_cost']} руб. с учетом НДС.",
        f"",
        "*в стоимость работ не включены:",
        "- проведение историко-культурной экспертизы участка (археологические изыскания);",
        "- дендрологические работы;",
        "- определение рыбохозяйственных характеристик и категорий водоемов, поиск зимовальных ям, согласование с Росрыболовством;",
        "- научное изучение растительного и животного мира, краснокнижных видов растений, грибов и животных, поиск мест гнездования птиц;",
        "- научное изучение водных биоресурсов водоемов;",
        "- биотестирование почво-грунтов с целью их экотоксикологической оценки и определения класса опасности отходов грунта."
    ]
    return "\n".join(lines)

def build_igmi_section(data):
    lines = [
        "Комплекс гидрометеорологических изысканий:",
        f"4.1. Полевые работы:",
        f"- рекогносцировочное обследование, кат. сложности II – {data['igmi_route_km']} км;",
        f"- фотоработы – {data['igmi_photo_count']} снимков.",
        f"4.2. Камеральные работы:",
        f"- рекогносцировочное обследование. Обработка полевых материалов – {data['igmi_route_km']} км;",
        f"- составление таблицы гидрологической изученности бассейна реки при числе пунктов наблюдений до 50 – 1 таблица;",
        f"- составление схемы гидрометеорологической изученности бассейна реки при числе пунктов наблюдений до 50 – 1 схема;",
        f"- подбор станций или постов с оценкой качества материалов наблюдений и степени их репрезентативности – 1 станция;",
        f"- построение розы ветров – {data['igmi_wind_rose_count']} графика;",
        f"- оценка возможности затопления/подтопления водными объектами участка изысканий – 1 оценка;",
        f"- составление климатической характеристики района – 1 записка;",
        f"- составление программы работ – 1 программа;",
        f"- составление технического отчета – 1 отчет.",
        f"",
        f"Сроки выполнения работ: {data['igmi_duration_days']} календарных дней.",
        f"Стоимость работ по п.4 составит: {data['igmi_cost']} руб. с учетом НДС.",
        f"*в стоимость работ не включены:",
        f"- полная климатическая справка - по требованию экспертизы."
    ]
    return "\n".join(lines)

@app.get("/generate-kp")
async def generate_kp(
    deal_id: str = "",
    object_name: str = "",
    address: str = "",
    cadastral_number: str = "",
    date: str = None,
    total_cost: str = "0",
    advance_percent: str = "50",
    validity_days: str = "30",
    igi: str = "0",
    igi_drilling_depth: str = "5",
    igi_boreholes: str = "4",
    igi_sounding_points: str = "4",
    igi_duration_days: str = "35",
    igi_cost: str = "0",
    igdi: str = "0",
    igdi_area_ha: str = "0",
    igdi_scale: str = "1:500",
    igdi_contour_interval: str = "0.5",
    igdi_duration_days: str = "50",
    igdi_survey_days: str = "15",
    igdi_coordination_days: str = "35",
    igdi_cost: str = "0",
    igdi_survey_cost: str = "0",
    igdi_coordination_cost: str = "0",
    igdi_report_cost: str = "0",
    iei: str = "0",
    iei_area_ha: str = "0",
    iei_gamma_points: str = "0",
    iei_noise_points: str = "0",
    iei_emi_points: str = "0",
    iei_soil_samples: str = "0",
    iei_bio_samples: str = "0",
    iei_rad_samples: str = "0",
    iei_surface_water_samples: str = "0",
    iei_sediment_samples: str = "0",
    iei_water_samples: str = "0",
    iei_water_boreholes: str = "0",
    iei_layered_samples_deep: str = "0",
    iei_deep_boreholes: str = "0",
    iei_layered_samples_shallow: str = "0",
    iei_shallow_boreholes: str = "0",
    iei_background_soil_samples: str = "0",
    iei_agro_samples: str = "0",
    iei_pits: str = "0",
    iei_duration_days: str = "35",
    iei_cost: str = "0",
    igmi: str = "0",
    igmi_route_km: str = "0",
    igmi_photo_count: str = "0",
    igmi_wind_rose_count: str = "0",
    igmi_duration_days: str = "40",
    igmi_cost: str = "0"
):
    try:
        # Подготовка данных
        data = {
            "object_name": object_name or "—",
            "address": address or "—",
            "cadastral_number": cadastral_number or "—",
            "date": date or datetime.now().strftime("%d.%m.%Y"),
            "total_cost": format_cost(total_cost),
            "advance_percent": advance_percent,
            "validity_days": validity_days,
            "igi_drilling_depth": igi_drilling_depth,
            "igi_boreholes": igi_boreholes,
            "igi_sounding_points": igi_sounding_points,
            "igi_duration_days": igi_duration_days,
            "igi_cost": format_cost(igi_cost),
            "igdi_area_ha": igdi_area_ha,
            "igdi_scale": igdi_scale,
            "igdi_contour_interval": igdi_contour_interval,
            "igdi_duration_days": igdi_duration_days,
            "igdi_survey_days": igdi_survey_days,
            "igdi_coordination_days": igdi_coordination_days,
            "igdi_cost": format_cost(igdi_cost),
            "igdi_survey_cost": format_cost(igdi_survey_cost),
            "igdi_coordination_cost": format_cost(igdi_coordination_cost),
            "igdi_report_cost": format_cost(igdi_report_cost),
            "iei_area_ha": iei_area_ha,
            "iei_gamma_points": iei_gamma_points,
            "iei_noise_points": iei_noise_points,
            "iei_emi_points": iei_emi_points,
            "iei_soil_samples": iei_soil_samples,
            "iei_bio_samples": iei_bio_samples,
            "iei_rad_samples": iei_rad_samples,
            "iei_surface_water_samples": iei_surface_water_samples,
            "iei_sediment_samples": iei_sediment_samples,
            "iei_water_samples": iei_water_samples,
            "iei_water_boreholes": iei_water_boreholes,
            "iei_layered_samples_deep": iei_layered_samples_deep,
            "iei_deep_boreholes": iei_deep_boreholes,
            "iei_layered_samples_shallow": iei_layered_samples_shallow,
            "iei_shallow_boreholes": iei_shallow_boreholes,
            "iei_background_soil_samples": iei_background_soil_samples,
            "iei_agro_samples": iei_agro_samples,
            "iei_pits": iei_pits,
            "iei_duration_days": iei_duration_days,
            "iei_cost": format_cost(iei_cost),
            "igmi_route_km": igmi_route_km,
            "igmi_photo_count": igmi_photo_count,
            "igmi_wind_rose_count": igmi_wind_rose_count,
            "igmi_duration_days": igmi_duration_days,
            "igmi_cost": format_cost(igmi_cost)
        }

        sections = []
        if igi == "1":
            sections.append(("ИГИ", build_igi_section(data)))
        if igdi == "1":
            sections.append(("ИГДИ", build_igdi_section(data)))
        if iei == "1":
            sections.append(("ИЭИ", build_iei_section(data)))
        if igmi == "1":
            sections.append(("ИГМИ", build_igmi_section(data)))

        content_lines = []
        for idx, (name, text) in enumerate(sections, start=1):
            prefix_map = {"ИГИ": "1.", "ИГДИ": "2.", "ИЭИ": "3.", "ИГМИ": "4."}
            old_prefix = prefix_map.get(name, "1.")
            new_text = text.replace(old_prefix, f"{idx}.")

            for old_num in ["1", "2", "3", "4"]:
                new_text = new_text.replace(f"п.{old_num}", f"п.{idx}")
                new_text = new_text.replace(f"п. {old_num}", f"п. {idx}")

            lines = new_text.split("\n")
            lines[0] = f"{idx}. {lines[0]}"
            content_lines.extend(lines)
            content_lines.append("")

        content = "\n".join(content_lines).strip()

        template_path = os.path.join(os.path.dirname(__file__), "templates", "kp_template.docx")
        if not os.path.exists(template_path):
            raise HTTPException(status_code=400, detail="Шаблон КП не найден")

        doc = Document(template_path)

        for paragraph in doc.paragraphs:
            text = paragraph.text
            if "{{content}}" in text:
                paragraph.text = text.replace("{{content}}", content)
            else:
                for key, value in data.items():
                    placeholder = f"{{{{{key}}}}}"
                    if placeholder in text:
                        text = text.replace(placeholder, str(value))
                paragraph.text = text

        safe_name = "".join(c for c in object_name if c.isalnum() or c in " _-")
        filename = f"KP_{safe_name}_{datetime.now().strftime('%Y%m%d')}.docx"
        output_path = f"/tmp/{filename}"
        doc.save(output_path)

        # Загрузка в Bitrix24
        async with httpx.AsyncClient(timeout=30) as client:
            prep_resp = await client.post(
                f"{WEBHOOK}disk.folder.uploadfile.json",
                data={"id": DISK_FOLDER_ID}
            )
            prep_data = prep_resp.json()
            if "result" not in prep_data or "uploadUrl" not in prep_data["result"]:
                raise HTTPException(status_code=500, detail="Не удалось получить uploadUrl от Bitrix24")

            upload_url = prep_data["result"]["uploadUrl"]
            field_name = prep_data["result"].get("field", "file")

            with open(output_path, "rb") as f:
                files = {field_name: (filename, f, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")}
                upload_resp = await client.post(upload_url, files=files)

            upload_result = upload_resp.json()
            if "result" not in upload_result:
                raise HTTPException(status_code=500, detail="Ошибка загрузки файла в Bitrix24")

            file_id = str(upload_result["result"]["ID"])

        download_url = f"https://izyskaniya.bitrix24.ru/disk/showFile/{file_id}/?filename={filename}"

        # Сохранение ссылки в сделку
        async with httpx.AsyncClient(timeout=30) as client:
            update_resp = await client.post(
    f"{WEBHOOK}crm.deal.update.json",
    json={
        "id": deal_id,
        "FIELDS": {  # ← ЗАГЛАВНЫМИ БУКВАМИ!
            KP_LINK_FIELD: download_url
        }
    }
)
            update_data = update_resp.json()
            if "result" not in update_data:
                raise HTTPException(status_code=500, detail="Не удалось сохранить ссылку в сделку")

        os.remove(output_path)

        return {
            "status": "success",
            "message": f"📄 КП готов! Скачать можно по ссылке: {download_url}",
            "download_url": download_url
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))