function byId(id) {
    return document.getElementById(id);
}

function updateStudentInfo(student) {
    byId("info_student_name").textContent = student.student_name || "-";
    byId("info_class_name").textContent = student.class_name || "-";
    byId("info_father_name").textContent = student.father_name || "-";
    byId("info_father_phone").textContent = student.father_phone || "-";
    byId("info_mother_name").textContent = student.mother_name || "-";
    byId("info_mother_phone").textContent = student.mother_phone || "-";
    byId("info_area").textContent = student.area || "-";
}

function resetStudentPreview() {
    byId("info_student_name").textContent = "-";
    byId("info_class_name").textContent = "-";
    byId("info_father_name").textContent = "-";
    byId("info_father_phone").textContent = "-";
    byId("info_mother_name").textContent = "-";
    byId("info_mother_phone").textContent = "-";
    byId("info_area").textContent = "-";
    byId("info_previous_count").textContent = "0";

    const actionBox = byId("actionBox");
    actionBox.textContent = "الإجراء التلقائي سيظهر هنا";
    actionBox.className = "action-box";

    const motherBtn = byId("motherWhatsappBtn");
    motherBtn.href = "#";
    motherBtn.classList.add("disabled-link");

    const fatherBtn = byId("fatherWhatsappBtn");
    fatherBtn.href = "#";
    fatherBtn.classList.add("disabled-link");
}

async function loadStudentOnly() {
    const studentId = byId("studentSelect").value;
    if (!studentId) {
        resetStudentPreview();
        return;
    }

    const res = await fetch(`/api/student/${studentId}`);
    const data = await res.json();

    if (data.ok) {
        updateStudentInfo(data.student);
    }
}

async function previewAction() {
    const studentId = byId("studentSelect").value;
    const violationId = byId("violationSelect").value;
    const note = byId("noteInput").value;

    if (!studentId) {
        resetStudentPreview();
        return;
    }

    await loadStudentOnly();

    if (!violationId) {
        byId("info_previous_count").textContent = "0";
        return;
    }

    const res = await fetch("/api/preview-action", {
        method: "POST",
        headers: {"Content-Type": "application/json", "X-CSRF-Token": document.querySelector('meta[name="csrf-token"]').content},
        body: JSON.stringify({
            student_id: studentId,
            violation_id: violationId,
            note: note
        })
    });

    const data = await res.json();

    if (!data.ok) {
        return;
    }

    updateStudentInfo(data.student);
    byId("info_previous_count").textContent = data.previous_count;

    const actionBox = byId("actionBox");
    actionBox.textContent = data.result_badge;
    actionBox.className = "action-box level-" + data.result_level.toLowerCase();

    const motherBtn = byId("motherWhatsappBtn");
    if (data.mother_link) {
        motherBtn.href = data.mother_link;
        motherBtn.classList.remove("disabled-link");
    } else {
        motherBtn.href = "#";
        motherBtn.classList.add("disabled-link");
    }

    const fatherBtn = byId("fatherWhatsappBtn");
    if (data.father_link) {
        fatherBtn.href = data.father_link;
        fatherBtn.classList.remove("disabled-link");
    } else {
        fatherBtn.href = "#";
        fatherBtn.classList.add("disabled-link");
    }
}

function editStudent(student) {
    byId("student_id").value = student.id || "";
    byId("student_name").value = student.student_name || "";
    byId("class_name").value = student.class_name || "";
    byId("father_name").value = student.father_name || "";
    byId("father_phone").value = student.father_phone || "";
    byId("mother_name").value = student.mother_name || "";
    byId("mother_phone").value = student.mother_phone || "";
    byId("area").value = student.area || "";

    location.hash = "#students-section";
}

function resetStudentForm() {
    byId("student_id").value = "";
    byId("student_name").value = "";
    byId("class_name").value = "";
    byId("father_name").value = "";
    byId("father_phone").value = "";
    byId("mother_name").value = "";
    byId("mother_phone").value = "";
    byId("area").value = "";
}

function editCatalog(item) {
    byId("violation_id").value = item.id || "";
    byId("violation_title").value = item.title || "";
    byId("base_level").value = item.base_level || "A";

    location.hash = "#catalog-section";
}

function resetCatalogForm() {
    byId("violation_id").value = "";
    byId("violation_title").value = "";
    byId("base_level").value = "A";
}

function setupStudentSearch() {
    const searchInput = byId("studentSearch");
    const select = byId("studentSelect");

    if (!searchInput || !select) return;

    searchInput.addEventListener("input", function () {
        const search = this.value.trim().toLowerCase();
        const options = select.options;

        for (let i = 0; i < options.length; i++) {
            if (i === 0) continue;
            const text = options[i].text.toLowerCase();
            options[i].hidden = !text.includes(search);
        }
    });
}

function setupTabs() {
    const links = document.querySelectorAll(".tab-link");
    links.forEach(link => {
        link.addEventListener("click", () => {
            links.forEach(l => l.classList.remove("active"));
            link.classList.add("active");
        });
    });
}

function initCharts() {
    const barCtx = document.getElementById("barChart");
    const pieCtx = document.getElementById("pieChart");
    const lineCtx = document.getElementById("lineChart");

    if (barCtx) {
        new Chart(barCtx, {
            type: "bar",
            data: {
                labels: levelKeys.map(k => levelLabels[k]),
                datasets: [{
                    label: "عدد السجلات حسب المستوى",
                    data: byLevel
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: { labels: { color: "#fff" } }
                },
                scales: {
                    x: { ticks: { color: "#fff" } },
                    y: { ticks: { color: "#fff" } }
                }
            }
        });
    }

    if (pieCtx) {
        new Chart(pieCtx, {
            type: "pie",
            data: {
                labels: violationNames,
                datasets: [{
                    label: "حسب نوع المخالفة",
                    data: violationCounts
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: { labels: { color: "#fff" } }
                }
            }
        });
    }

    if (lineCtx) {
        new Chart(lineCtx, {
            type: "line",
            data: {
                labels: months,
                datasets: [{
                    label: "الخط الزمني الشهري",
                    data: monthlyCounts,
                    tension: 0.3
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: { labels: { color: "#fff" } }
                },
                scales: {
                    x: { ticks: { color: "#fff" } },
                    y: { ticks: { color: "#fff" } }
                }
            }
        });
    }
}

document.addEventListener("DOMContentLoaded", function () {
    setupStudentSearch();
    setupTabs();
    initCharts();

    const studentSelect = byId("studentSelect");
    const violationSelect = byId("violationSelect");
    const noteInput = byId("noteInput");

    if (studentSelect) studentSelect.addEventListener("change", previewAction);
    if (violationSelect) violationSelect.addEventListener("change", previewAction);
    if (noteInput) noteInput.addEventListener("input", previewAction);
});