# app.py (final version)
import streamlit as st
from solver import solve, get_available_courses, load_lectures, load_indexes, analyze_infeasibility

st.set_page_config(page_title="NTU Timetable Optimizer", layout="centered")
st.title("🎓 NTU Timetable Optimizer")
st.caption("Minimize campus days • Avoid conflicts • Powered by OR-Tools")

@st.cache_resource
def load_courses():
    return get_available_courses()

all_courses = load_courses()

selected = st.multiselect(
    "Select your courses",
    options=all_courses,
    default=["SC2079", "SC2104", "SC3103", "SC3021"]
)

if st.button("🔍 Optimize Timetable"):
    if not selected:
        st.warning("Please select at least one course.")
    else:
        with st.spinner("Solving... (5–10 seconds)"):
            lectures = load_lectures("Table1.xlsx", selected)
            index_map = load_indexes("Table2.xlsx", selected, "all")
            result = solve(lectures, index_map, selected)

        if result["success"]:
            st.success(f"✅ Optimized! **{result['campus_days']} campus days**")
            
            st.subheader("📌 Selected Indexes")
            for course, idx in result["indexes"].items():
                st.write(f"- **{course}**: Index `{idx}`")
            
            st.subheader("📅 Weekly Timetable")
            days = ["Mon", "Tue", "Wed", "Thu", "Fri"]
            # Group by day
            for day_idx, day_name in enumerate(days):
                day_sessions = [s for s in result["timetable"] if s["day"] == day_idx]
                if day_sessions:
                    st.markdown(f"**{day_name}**")
                    for s in sorted(day_sessions, key=lambda x: x["start"]):
                        start_h = s["start"] // 60
                        start_m = s["start"] % 60
                        end_h = s["end"] // 60
                        end_m = s["end"] % 60
                        weeks = " (all weeks)" if s["weeks"] == set(range(1,14)) else f" (Weeks: {sorted(s['weeks'])})"
                        st.text(f"  {start_h:02d}:{start_m:02d}–{end_h:02d}:{end_m:02d} | {s['course']} {s['type']}{weeks}")
        else:
            # Display the conflict message from the solver
            st.error(f"❌ {result['error']}")