import pandas as pd
from ortools.sat.python import cp_model

# Try to import tkinter, but handle if it's not available
try:
    import tkinter as tk
    from tkinter import ttk, messagebox
    TK_AVAILABLE = True
except ImportError:
    TK_AVAILABLE = False
    print("GUI not available. Using console input instead.")


# ============================================================
# Helper functions
# ============================================================

def time_to_minutes(t):
    """Convert HH:MM -> integer minutes."""
    if isinstance(t, str):
        t = t.strip()
        # Handle cases where time might have extra spaces or different formats
        if ":" in t:
            h, m = map(int, t.split(":"))
        else:
            # Handle cases like "900" -> "09:00"
            if len(t) <= 2:
                h, m = int(t), 0
            elif len(t) == 3:
                h, m = int(t[0]), int(t[1:])
            elif len(t) == 4:
                h, m = int(t[:2]), int(t[2:])
            else:
                raise ValueError(f"Cannot parse time: {t}")
        return h * 60 + m
    return int(t)


def day_to_index(day):
    """Map weekday strings to integer indices."""
    # Convert to title case to match the mapping
    day = day.strip().title()
    
    mapping = {
        "Mon": 0, "Monday": 0,
        "Tue": 1, "Tuesday": 1,
        "Wed": 2, "Wednesday": 2,
        "Thu": 3, "Thursday": 3,
        "Fri": 4, "Friday": 4
    }
    # Handle case where day might not be in mapping
    if day not in mapping:
        raise ValueError(f"Unknown day: {day}")
    return mapping[day]


def parse_weeks(remark, week_pattern="all"):
    """Parse 'Teaching Wk2,4,6' style strings into sets with week pattern filtering."""
    # Default weeks based on pattern
    if week_pattern == "even":
        default_weeks = {2, 4, 6, 8, 10, 12}
    elif week_pattern == "odd":
        default_weeks = {1, 3, 5, 7, 9, 11, 13}
    else:
        default_weeks = set(range(1, 14))  # Weeks 1-13 default
    
    if not isinstance(remark, str) or remark.strip() == "":
        return default_weeks

    # Extract numbers
    remark = remark.replace("Teaching Wk", "")
    week_nums = remark.split(",")
    try:
        parsed_weeks = set(int(w.strip()) for w in week_nums)
        # Apply week pattern filter
        if week_pattern == "even":
            return parsed_weeks.intersection({2, 4, 6, 8, 10, 12})
        elif week_pattern == "odd":
            return parsed_weeks.intersection({1, 3, 5, 7, 9, 11, 13})
        else:
            return parsed_weeks
    except ValueError as e:
        print(f"Warning: Could not parse weeks from '{remark}': {e}")
        return default_weeks

# ============================================================
# Load Data
# ============================================================

def load_lectures(filepath, selected_courses):
    df = pd.read_excel(filepath)

    sessions = []

    for _, row in df.iterrows():
        if row["Course Code"] not in selected_courses:
            continue
            
        try:
            # Skip rows with missing time data (courses with only labs)
            if pd.isna(row["Start Time"]) or pd.isna(row["End Time"]):
                print(f"Skipping lecture for {row['Course Code']} - no time data (likely lab-only course)")
                continue
                
            sessions.append({
                "course": row["Course Code"],
                "type": row["TYPE"],
                "day": day_to_index(row["Day"]),
                "start": time_to_minutes(row["Start Time"]),
                "end": time_to_minutes(row["End Time"]),
                "weeks": set(range(1, 14))  # lectures always every week
            })
        except Exception as e:
            print(f"Error processing lecture row: {row['Course Code']} - {e}")
            continue

    return sessions


def load_indexes(filepath, selected_courses, week_pattern="all"):
    df = pd.read_excel(filepath)

    index_map = {}

    for _, row in df.iterrows():
        course = row["Course Code"]
        if course not in selected_courses:
            continue

        idx = row["Index"]

        if (course, idx) not in index_map:
            index_map[(course, idx)] = []

        try:
            index_map[(course, idx)].append({
                "course": course,
                "index": idx,
                "type": row["TYPE"],  # TUT or LAB
                "day": day_to_index(row["Day"]),
                "start": time_to_minutes(row["Start Time"]),
                "end": time_to_minutes(row["End Time"]),
                "weeks": parse_weeks(row.get("Remark", ""), week_pattern)
            })
        except Exception as e:
            print(f"Error processing index row: {row['Course Code']} - {e}")
            continue

    return index_map


# ============================================================
# Overlap detection helpers
# ============================================================

def sessions_overlap(a, b):
    """Check if two time intervals overlap (ignores week logic)."""
    if a["day"] != b["day"]:
        return False
    return not (a["end"] <= b["start"] or b["end"] <= a["start"])


def strict_conflict(a, b):
    """
    TUT conflicts with ANYTHING if overlapping.
    LAB conflicts with:
      - lectures
      - tutorials
      - labs whose weeks intersect
    """
    # If sessions do not overlap in time OR day, no issue.
    if not sessions_overlap(a, b):
        return False
        
    # Special case: LAB vs LAB with non-overlapping weeks is allowed
    if a["type"] == "LAB" and b["type"] == "LAB":
        return labs_conflict(a, b)
        
    # All other overlapping sessions conflict
    return True


def labs_conflict(a, b):
    """
    LAB-specific rule:
    Overlap allowed ONLY if their weeks do NOT intersect.
    """
    # If sessions do not overlap in time OR day, no issue.
    if not sessions_overlap(a, b):
        return False

    # Check weekly intersection - if weeks overlap, they conflict
    return bool(a["weeks"].intersection(b["weeks"]))

# ============================================================
# Build OR-Tools Model
# ============================================================

def build_model(lectures, index_map, selected_courses):
    model = cp_model.CpModel()

    # --------------------------------------
    # Decision variables: choose one index per course
    # --------------------------------------
    chosen = {}  # (course, index) -> BoolVar

    for (course, idx) in index_map:
        chosen[(course, idx)] = model.NewBoolVar(f"{course}_{idx}")

    # Each course must choose exactly ONE index
    for course in selected_courses:
        eligible = [chosen[(c, idx)] for (c, idx) in index_map if c == course]
        if eligible:  # Only add constraint if there are eligible indexes
            model.Add(sum(eligible) == 1)
        else:
            print(f"Warning: No indexes found for course {course}")

    # --------------------------------------
    # Conflict constraints - Apply ALL valid constraints
    # --------------------------------------
    all_index_sessions = []
    for key, sessions in index_map.items():
        (course, idx) = key
        for s in sessions:
            all_index_sessions.append((course, idx, s))

    # Apply lecture conflict constraints - Apply ALL valid constraints
    for i in range(len(all_index_sessions)):
        c1, i1, s1 = all_index_sessions[i]

        # Check vs lectures - Apply ALL valid constraints
        for lec in lectures:
            if strict_conflict(s1, lec):
                # Check if this specific session conflicts with the lecture in any week
                # If so, this index cannot be chosen
                if s1["weeks"].intersection(lec["weeks"]):
                    model.Add(chosen[(c1, i1)] == 0)

    # Apply index vs index conflict constraints - Apply ALL valid constraints
    conflict_count = 0
    
    for i in range(len(all_index_sessions)):
        c1, i1, s1 = all_index_sessions[i]

        for j in range(i + 1, len(all_index_sessions)):
            c2, i2, s2 = all_index_sessions[j]

            # Only check sessions from different courses
            if c1 == c2:
                continue

            if strict_conflict(s1, s2):
                # Check if they actually conflict in weeks
                if s1["weeks"].intersection(s2["weeks"]):
                    # Apply the constraint to prevent both indexes from being selected together
                    model.Add(chosen[(c1, i1)] + chosen[(c2, i2)] <= 1)
                    conflict_count += 1

    # --------------------------------------
    # Campus day variables
    # --------------------------------------
    campus_day = [model.NewBoolVar(f"campus_{d}") for d in range(5)]

    # A day is campus day if any chosen session (TUT or LAB) is on that day
    # Lectures do NOT count as campus days
    for d in range(5):
        day_sessions = []
        for (course, idx), sessions in index_map.items():
            for s in sessions:
                # Only TUT/LAB sessions count for campus days
                if s["day"] == d and (s["type"] == "TUT" or s["type"] == "LAB"):
                    day_sessions.append(chosen[(course, idx)])

        if day_sessions:
            model.AddMaxEquality(campus_day[d], day_sessions)
        else:
            model.Add(campus_day[d] == 0)

    # --------------------------------------
    # Objective: Minimize campus days
    # --------------------------------------
    model.Minimize(sum(campus_day))

    return model, chosen, campus_day


# ============================================================
# Solve wrapper
# ============================================================

def solve(lectures, index_map, selected_courses, num_solutions=3):
    model, chosen, campus_day = build_model(lectures, index_map, selected_courses)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 10
    
    # Store multiple solutions
    solutions = []
    
    # Callback to collect solutions
    class SolutionCollector(cp_model.CpSolverSolutionCallback):
        def __init__(self, chosen, campus_day, max_solutions):
            cp_model.CpSolverSolutionCallback.__init__(self)
            self.chosen = chosen
            self.campus_day = campus_day
            self.solutions = []
            self.max_solutions = max_solutions
            self.solution_count = 0
            
        def on_solution_callback(self):
            if self.solution_count >= self.max_solutions:
                self.StopSearch()
                return
                
            # Store the current solution
            solution = {
                "chosen": {},
                "campus_days": sum(self.Value(campus_day[d]) for d in range(5))
            }
            
            for (course, idx), var in self.chosen.items():
                solution["chosen"][(course, idx)] = self.Value(var)
                
            self.solutions.append(solution)
            self.solution_count += 1
    
    # Create solution collector
    collector = SolutionCollector(chosen, campus_day, num_solutions)
    
    # Enable solution enumeration
    # Note: For problems with tight constraints, there may be only one optimal solution
    # We'll collect multiple feasible solutions if available
    solver.parameters.enumerate_all_solutions = True
    
    # Solve with solution collection
    result = solver.Solve(model, collector)
    
    print("\nSolver status:", solver.StatusName(result))
    
    # Use collected solutions
    solutions = collector.solutions
    
    if not solutions:
        print("No feasible solution found!")
        analyze_infeasibility(lectures, index_map, selected_courses)
        return None
    
    # Sort solutions by number of campus days (ascending)
    solutions.sort(key=lambda s: s["campus_days"])
    
    # --------------------------------------
    # Output multiple solutions
    # --------------------------------------
    num_found = min(num_solutions, len(solutions))
    if num_found == 1:
        print(f"\n=== Found 1 Optimal Solution ===")
    else:
        print(f"\n=== Top {num_found} Optimal Solutions ===")
    
    for i, solution in enumerate(solutions[:num_solutions]):
        print(f"\n--- Solution #{i+1} ({solution['campus_days']} campus days) ---")
        for (course, idx), val in solution["chosen"].items():
            if val == 1:
                print(f"  {course}: Index {idx}")
    
    # Only show note if we found fewer solutions than requested
    if num_found < num_solutions and num_solutions > 1:
        print(f"\n(Note: Only {num_found} feasible solution(s) found for these courses)")
    
    return solutions[:num_solutions]

def check_solution_conflicts(lectures, index_map, chosen, solver):
    """Check for conflicts in the current solution."""
    conflicts = []
    
    # Get all selected sessions
    selected_sessions = []
    
    # Add lectures
    for lec in lectures:
        selected_sessions.append({
            "course": lec["course"],
            "type": lec["type"],
            "day": lec["day"],
            "start": lec["start"],
            "end": lec["end"],
            "weeks": lec["weeks"]
        })
    
    # Add selected index sessions
    for (course, idx), var in chosen.items():
        if solver.Value(var) == 1:
            sessions = index_map[(course, idx)]
            for session in sessions:
                selected_sessions.append({
                    "course": session["course"],
                    "type": session["type"],
                    "day": session["day"],
                    "start": session["start"],
                    "end": session["end"],
                    "weeks": session["weeks"]
                })
    
    # Check all pairs for conflicts
    for i in range(len(selected_sessions)):
        for j in range(i + 1, len(selected_sessions)):
            s1 = selected_sessions[i]
            s2 = selected_sessions[j]
            
            # Check if they overlap in time and day
            if s1["day"] == s2["day"] and sessions_overlap(s1, s2):
                # Check if they actually conflict based on our rules
                if strict_conflict(s1, s2):
                    # Check if they actually conflict in weeks
                    if s1["weeks"].intersection(s2["weeks"]):
                        conflicts.append(
                            f"{s1['course']} {s1['type']} conflicts with {s2['course']} {s2['type']} "
                            f"on day {s1['day']} at {s1['start']//60:02d}:{s1['start']%60:02d}-"
                            f"{s1['end']//60:02d}:{s1['end']%60:02d} vs "
                            f"{s2['start']//60:02d}:{s2['start']%60:02d}-{s2['end']//60:02d}:{s2['end']%60:02d} "
                            f"(weeks: {sorted(s1['weeks'].intersection(s2['weeks']))})"
                        )
    
    return conflicts

def show_detailed_timetable(lectures, index_map, chosen, solver):
    """Display detailed timetable information."""
    # Collect all sessions for display
    timetable_sessions = []
    
    # Add lecture sessions
    for lec in lectures:
        timetable_sessions.append({
            "course": lec["course"],
            "type": lec["type"],
            "day": lec["day"],
            "start": lec["start"],
            "end": lec["end"],
            "weeks": lec["weeks"]
        })
    
    # Add selected index sessions
    for (course, idx), var in chosen.items():
        if solver.Value(var) == 1:
            sessions = index_map[(course, idx)]
            for session in sessions:
                timetable_sessions.append({
                    "course": session["course"],
                    "type": session["type"],
                    "day": session["day"],
                    "start": session["start"],
                    "end": session["end"],
                    "weeks": session["weeks"]
                })
    
    # Sort by day and start time
    timetable_sessions.sort(key=lambda x: (x["day"], x["start"]))
    
    # Display timetable
    days = ["Mon", "Tue", "Wed", "Thu", "Fri"]
    for day_idx in range(5):
        day_sessions = [s for s in timetable_sessions if s["day"] == day_idx]
        if day_sessions:
            print(f"\n{days[day_idx]}:")
            for session in day_sessions:
                start_time = f"{session['start'] // 60:02d}:{session['start'] % 60:02d}"
                end_time = f"{session['end'] // 60:02d}:{session['end'] % 60:02d}"
                weeks_str = ""
                if session["weeks"] != set(range(1, 14)):
                    weeks_str = f" (Weeks: {sorted(session['weeks'])})"
                print(f"  {start_time}-{end_time} {session['course']} {session['type']}{weeks_str}")

def calculate_campus_days_for_weeks(lectures, index_map, chosen, solver, target_weeks):
    """Calculate campus days for a specific set of weeks using the current solution."""
    # Collect all sessions that occur in the target weeks
    campus_days = set()
    
    # Add selected index sessions that occur in target weeks
    # (Only TUT/LAB sessions count for campus days, not lectures)
    for (course, idx), var in chosen.items():
        if solver.Value(var) == 1:
            sessions = index_map[(course, idx)]
            for session in sessions:
                if session["weeks"].intersection(target_weeks):
                    campus_days.add(session["day"])
    
    return len(campus_days)

def analyze_infeasibility(lectures, index_map, selected_courses):
    """Analyze why a solution is infeasible."""
    print("\n=== Infeasibility Analysis ===")
    
    # Show just the first major conflict to keep it simple
    for course in selected_courses:
        # Get lectures for this course
        course_lectures = [lec for lec in lectures if lec["course"] == course]
        # Get indexes for other courses
        other_courses = [c for c in selected_courses if c != course]
        
        for lec in course_lectures:
            for other_course in other_courses:
                course_indexes = [(c, idx) for (c, idx) in index_map.keys() if c == other_course]
                for (c, idx) in course_indexes:
                    sessions = index_map[(c, idx)]
                    for session in sessions:
                        if strict_conflict(lec, session):
                            if lec["weeks"].intersection(session["weeks"]):
                                print(f"❌ Conflict: {course} lecture clashes with {other_course} {session['type']} (Index {idx})")
                                return  # Just show the first conflict
    
    # If no lecture conflicts, check index vs index conflicts
    course_pairs = [(c1, c2) for i, c1 in enumerate(selected_courses) for c2 in selected_courses[i+1:]]
    for c1, c2 in course_pairs:
        indexes1 = [(c, idx) for (c, idx) in index_map.keys() if c == c1]
        indexes2 = [(c, idx) for (c, idx) in index_map.keys() if c == c2]
        
        for (_, idx1) in indexes1:
            for (_, idx2) in indexes2:
                sessions1 = index_map[(c1, idx1)]
                sessions2 = index_map[(c2, idx2)]
                
                for s1 in sessions1:
                    for s2 in sessions2:
                        if strict_conflict(s1, s2):
                            if s1["weeks"].intersection(s2["weeks"]):
                                print(f"❌ Conflict: {c1} Index {idx1} clashes with {c2} Index {idx2}")
                                return  # Just show the first conflict
    
    print("❌ No feasible schedule found due to timing conflicts")

# ============================================================
# GUI/console for course selection
# ============================================================

def select_courses(available_courses):
    """Select courses using GUI if available, otherwise console input."""
    print(f"Available courses: {sorted(available_courses)}")
    if TK_AVAILABLE:
        print("GUI is available, attempting to show course selection dialog...")
        return select_courses_gui(available_courses)
    else:
        print("GUI not available, using console input...")
        return select_courses_console(available_courses)

def select_courses_gui(available_courses):
    """Create a GUI popup for users to select courses."""
    try:
        print("Creating Tkinter root window...")
        root = tk.Tk()
        root.title("Course Selection")
        root.geometry("400x500")
        root.resizable(True, True)
        
        # Store selected courses
        selected_courses = []
        
        # Create a frame for the course list
        frame = ttk.Frame(root, padding="10")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(frame, text="Select Courses to Register For:", font=("Arial", 12, "bold"))
        title_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        # Instructions
        instruction_label = ttk.Label(frame, text="Select at least 2 courses:")
        instruction_label.grid(row=1, column=0, sticky=tk.W, pady=(0, 10))
        
        # Create a canvas and scrollbar for the course list
        canvas = tk.Canvas(frame)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Create checkboxes for each course (initially unchecked)
        course_vars = {}
        sorted_courses = sorted(available_courses)
        for i, course in enumerate(sorted_courses):
            var = tk.BooleanVar(value=False)  # Explicitly set to False
            course_vars[course] = var
            checkbox = ttk.Checkbutton(scrollable_frame, text=course, variable=var)
            checkbox.grid(row=i, column=0, sticky=tk.W, padx=5, pady=2)
        
        canvas.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=2, column=1, sticky=(tk.N, tk.S))
        
        # Button frame
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=10)
        
        def submit_selection():
            selected = [course for course, var in course_vars.items() if var.get()]
            print(f"User selected courses: {selected}")
            if len(selected) < 2:
                messagebox.showwarning("Insufficient Selection", "Please select at least 2 courses.")
                return
            selected_courses.extend(selected)
            root.destroy()
        
        def select_all():
            for var in course_vars.values():
                var.set(True)
        
        def deselect_all():
            for var in course_vars.values():
                var.set(False)
        
        # Buttons
        submit_btn = ttk.Button(button_frame, text="Generate Timetable", command=submit_selection)
        submit_btn.pack(side=tk.RIGHT, padx=5)
        
        select_all_btn = ttk.Button(button_frame, text="Select All", command=select_all)
        select_all_btn.pack(side=tk.LEFT, padx=5)
        
        deselect_all_btn = ttk.Button(button_frame, text="Deselect All", command=deselect_all)
        deselect_all_btn.pack(side=tk.LEFT, padx=5)
        
        # Handle window closing
        def on_closing():
            print("Window closing...")
            if messagebox.askokcancel("Quit", "Do you want to quit without generating a timetable?"):
                root.destroy()
                selected_courses.clear()  # Clear to indicate cancellation
        
        root.protocol("WM_DELETE_WINDOW", on_closing)
        
        # Center the window
        root.update_idletasks()
        x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
        y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
        root.geometry(f"+{x}+{y}")
        
        print("Starting mainloop... Please interact with the GUI window.")
        root.mainloop()
        print(f"Mainloop ended. Selected courses: {selected_courses}")
        
        return selected_courses if selected_courses else None
    except Exception as e:
        print(f"GUI error: {e}. Falling back to console input.")
        return select_courses_console(available_courses)

def select_courses_console(available_courses):
    """Select courses using console input."""
    print("\n=== Course Selection ===")
    print("Available courses:")
    sorted_courses = sorted(available_courses)
    for i, course in enumerate(sorted_courses, 1):
        print(f"{i:2d}. {course}")
    
    print(f"\nPlease select courses by entering numbers separated by commas.")
    print(f"Example: 1,3,5")
    print(f"Or type 'all' to select all courses.")
    print(f"You must select at least 2 courses.\n")
    
    while True:
        try:
            user_input = input("Enter your selection: ").strip()
            
            if user_input.lower() == 'all':
                return sorted_courses, "all"
            
            if not user_input:
                print("Please enter at least one course number.")
                continue
                
            # Parse the input
            selected_indices = []
            for part in user_input.split(','):
                num = int(part.strip())
                if 1 <= num <= len(sorted_courses):
                    selected_indices.append(num)
                else:
                    print(f"Invalid course number: {num}. Please enter numbers between 1 and {len(sorted_courses)}.")
                    raise ValueError("Invalid course number")
            
            # Convert indices to course codes
            selected_courses = [sorted_courses[idx - 1] for idx in selected_indices]
            
            # Remove duplicates while preserving order
            seen = set()
            unique_selected = []
            for course in selected_courses:
                if course not in seen:
                    seen.add(course)
                    unique_selected.append(course)
            
            if len(unique_selected) < 2:
                print("Please select at least 2 courses.")
                continue
                
            return unique_selected, "all"
            
        except ValueError:
            print("Invalid input. Please enter numbers separated by commas or 'all'.")
        except KeyboardInterrupt:
            print("\nOperation cancelled.")
            return None, None
        except Exception as e:
            print(f"Error: {e}. Please try again.")

# ============================================================
# Load all available courses
# ============================================================

def get_available_courses():
    """Extract all available courses from the Excel files."""
    try:
        # Load lectures
        df1 = pd.read_excel("Table1.xlsx")
        lecture_courses = set(str(code).strip() for code in df1["Course Code"].tolist() if pd.notna(code))
        
        # Load indexes
        df2 = pd.read_excel("Table2.xlsx")
        index_courses = set(str(code).strip() for code in df2["Course Code"].tolist() if pd.notna(code))
        
        # Combine and return all courses
        all_courses = lecture_courses.union(index_courses)
        # Filter out any non-course entries
        valid_courses = {course for course in all_courses if course and not course.startswith('Unnamed')}
        return sorted(list(valid_courses))
    except Exception as e:
        print(f"Error loading course data: {e}")
        return []

# ============================================================
# Main runner
# ============================================================

if __name__ == "__main__":
    available_courses = get_available_courses()
    if not available_courses:
        print("No courses available. Please check your Excel files.")
    else:
        print(f"Found {len(available_courses)} available courses.")
        selection_result = select_courses_console(available_courses)
        if selection_result and selection_result[0]:
            selected_courses, week_pattern = selection_result
                
            print(f"\nGenerating timetable for: {', '.join(selected_courses)}")
            print("Loading data...")
            lectures = load_lectures("Table1.xlsx", selected_courses)
            index_map = load_indexes("Table2.xlsx", selected_courses, "all")

            print("Solving...")
            solve(lectures, index_map, selected_courses, num_solutions=3)
