        
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = 30.0
        if solver.Solve(model) in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            output_data = []
            for item in clean_assignments:
                for d in range(num_days):
                    for p in range(num_periods):
                        if solver.Value(schedule[(item["idx"], item["c"], item["s"], item["t"], item["r"], d, p)]) == 1:
                            output_data.append({"الفصل": item["c"], "المادة": item["s"], "المدرس": item["t"], "القاعة": item["r"], "اليوم": days[d], "الحصة": f"الحصة {p + 1}"})
            
            df_result = pd.DataFrame(output_data)
            
            # Master Table Creation with Fixed Logic
            wb_master = Workbook()
            ws_master = wb_master.active
            ws_master.title = "الحصص_الشامل"
            # ... [Logic to build headers] ...
            
            # --- الجزء المُعدل للجدول الشامل ---
            for idx, cls in enumerate(sorted(list(classes)), start=5):
                ws_master.cell(row=idx, column=2, value=str(cls))
                col_cursor = 3
                for day in days:
                    for p in periods:
                        match = df_result[(df_result["الفصل"].astype(str).str.strip() == str(cls).strip()) & 
                                          (df_result["اليوم"].astype(str).str.strip() == str(day).strip()) & 
                                          (df_result["الحصة"].astype(str).str.strip() == str(p).strip())]
                        
                        if not match.empty:
                            cell_val = f"{match.iloc[0]['المادة']}\n({match.iloc[0].get('المدرس', '')})"
                        else:
                            cell_val = "متاحة"
                        ws_master.cell(row=idx, column=col_cursor, value=cell_val)
                        col_cursor += 1
            # ---------------------------------
            wb_master.save("all_classes_master_table.xlsx")
            df_result.to_excel("final_timetable.xlsx", index=False)
            st.session_state.generated = True
    except Exception as e:
        st.error(f"خطأ: {e}")

if st.session_state.generated:
    st.success("تم التوليد بنجاح!")
    with open("final_timetable.xlsx", "rb") as f: st.download_button("تحميل الجداول", f, "final_timetable.xlsx")
    with open("all_classes_master_table.xlsx", "rb") as f: st.download_button("تحميل الجدول الشامل", f, "all_classes_master_table.xlsx")
