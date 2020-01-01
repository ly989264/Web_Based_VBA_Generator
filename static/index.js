function addCondition() {
    var conditions = document.getElementById("conditions");
    var current_num = parseInt(conditions.getAttribute("number"));
    conditions.setAttribute("number", current_num+1);
    var new_dir = document.createElement("div");
    new_dir.innerHTML = '<h1>Job ' + (current_num + 1) + '</h1>' +
        '<button class="show_collapse" onclick="show_collapse(this)">Show/Collapse</button><br>';
    new_dir.classList.add("OP");
    var new_new_dir = document.createElement("div");
    new_dir.appendChild(new_new_dir);
    var worksheet_configuration = document.createElement("div");
    worksheet_configuration.classList.add("submode");
    worksheet_configuration.innerHTML = '<h2>Worksheet configuration:</h2>' +
        'current workbook: <input type="text" class="current_workbook" placeholder="null"><br>' +
        'target_workbook: <input type="text" class="target_workbook" placeholder="null"><br>' +
        'current_worksheet: <input type="text" class="current_worksheet"><br>' +
        'target_worksheet: <input type="text" class="target_worksheet"><br>';
    new_new_dir.appendChild(worksheet_configuration);
    var filter = document.createElement("div");
    filter.classList.add("submode");
    filter.innerHTML = '<h2>Filter configuration:</h2>' +
        '<button id="add_filters" onclick="addFilter(this.nextSibling.nextSibling)">Add Filters</button><br>' +
        '<div class="filters" id="filters"></div>';
    new_new_dir.appendChild(filter);
    var copy_option = document.createElement("div");
    copy_option.classList.add("submode");
    copy_option.innerHTML = '<h2>Copy options:</h2>';
    var col2col = document.createElement("div");
    var fix2col = document.createElement("div");
    col2col.innerHTML = 'Column-to-column relation:<br>' +
        '<button onclick="addCopyRelation(this.nextSibling.nextSibling)">Add new col2col relation</button><br>' +
        '<div class="col2cols"></div>';
    fix2col.innerHTML = 'Fix content-to-column relation:<br>' +
        '<button onclick="addCopyRelation(this.nextSibling.nextSibling)">Add new fix2col relation</button><br>' +
        '<div class="fix2cols"></div>';
    copy_option.appendChild(col2col);
    copy_option.appendChild(fix2col);
    new_new_dir.appendChild(copy_option);
    var options = document.createElement("div");
    options.classList.add("submode");
    options.innerHTML = '<h2>Options:</h2>' +
        'Append: <input type="checkbox" class="checkbox_append_true" value="true">true ' +
        '<input type="checkbox" class="checkbox_append_false" value="false">false<br>' +
        'Disable screen updating: <input type="checkbox" class="checkbox_screenupdating_true" value="true">true ' +
        '<input type="checkbox" class="checkbox_screenupdating_false" value="false">false<br>'
    new_new_dir.appendChild(options);
    conditions.appendChild(new_dir);
}

function addFilter(filters) {
    console.log(filters);
    var already_num = filters.childNodes.length;
    var new_filter = document.createElement("div");
    new_filter.classList.add("filter_div");
    filters.appendChild(new_filter);
    new_filter.innerHTML = '<h4>Filter No.' + (already_num + 1) + '</h4><br>' +
        'column (string format): <input type="text" class="filter_column"><br>' +
        'operator: <input type="text" class="filter_operator"><br>' +
        'criteria: <button class="add_criteria" onclick="addCriteria(this.nextSibling.nextSibling)">Add Criteria</button><br>' +
        '<div class="criterias" id="criterias"></div><br>';
}

function addCriteria(criterias) {
    console.log("Adding criteria...");
    var new_criteria = document.createElement("input");
    new_criteria.setAttribute("type", "text");
    new_criteria.classList.add("filter_criteria");
    criterias.appendChild(new_criteria);
}

function addCopyRelation(criteria) {
    console.log("Adding copy relation...");
    var new_div = document.createElement("div");
    new_div.classList.add("copy_relation_div");
    new_div.innerHTML = '<input type="text" class="copy_in1">   =>   <input type="text" class="copy_in3">  (Target => Source)';
    criteria.appendChild(new_div);
}

function show_collapse(current_node) {
    var target_node = current_node.nextSibling.nextSibling;
    if (target_node.hasAttribute("hidden")) {
        target_node.removeAttribute("hidden");
    } else {
        target_node.setAttribute("hidden", true);
    }
}

function collect() {
    // return an object
    var conditions = document.getElementById("conditions");
    var num_jobs = conditions.getElementsByClassName("OP").length;
    var final_result = {};
    console.log(num_jobs);
    var job_divs = conditions.getElementsByClassName("OP");
    let each;
    for (each of job_divs) {
        console.log(each);
        if ("job" in final_result) {
            // do nothing
        } else {
            final_result["job"] = new Array();
        }
        var result = {};
        var job_divs = each.getElementsByTagName("div")[0];
        let each_submode;
        for (each_submode of job_divs.getElementsByClassName("submode")) {
            console.log(each_submode);
            var submode_name = each_submode.getElementsByTagName("h2")[0].innerHTML;
            console.log(submode_name);
            if (submode_name.startsWith("Worksheet")) {
                // worksheet configuration
                console.log("worksheet configuration");
                result["worksheet_configuration"] = {};
                if (each_submode.getElementsByClassName("current_workbook")[0].value === "" || each_submode.getElementsByClassName("current_workbook")[0].value === "null") {
                    result["worksheet_configuration"]["current_workbook"] = null;
                } else {
                    result["worksheet_configuration"]["current_workbook"] = each_submode.getElementsByClassName("current_workbook")[0].value;
                }
                if (each_submode.getElementsByClassName("target_workbook")[0].value === "" || each_submode.getElementsByClassName("target_workbook")[0].value === "null") {
                    result["worksheet_configuration"]["target_workbook"] = null;
                } else {
                    result["worksheet_configuration"]["target_workbook"] = each_submode.getElementsByClassName("target_workbook")[0].value;
                }
                if (each_submode.getElementsByClassName("current_worksheet")[0].value === "") {
                    alert("Oops, seems like the input box of current worksheet is empty!");
                    return null;
                }
                result["worksheet_configuration"]["current_worksheet"] = each_submode.getElementsByClassName("current_worksheet")[0].value;
                if (each_submode.getElementsByClassName("target_worksheet")[0].value === "") {
                    alert("Oops, seems like the input box of target worksheet is empty!");
                    return null;
                }
                result["worksheet_configuration"]["target_worksheet"] = each_submode.getElementsByClassName("target_worksheet")[0].value;
            } else if (submode_name.startsWith("Filter")) {
                // filter
                console.log("filter configuration");
                result["filter"] = {};
                result["filter"]["filters"] = new Array();
                var submode_filters = each_submode.getElementsByClassName("filters")[0];
                let each_submode_filter;
                for (each_submode_filter of submode_filters.childNodes) {
                    var temp_filter = {};
                    temp_filter["column_str"] = each_submode_filter.getElementsByClassName("filter_column")[0].value.toUpperCase();
                    temp_filter["operator"] = each_submode_filter.getElementsByClassName("filter_operator")[0].value.toUpperCase();
                    temp_filter["criteria"] = new Array();
                    var criteria_div = each_submode_filter.getElementsByClassName("criterias")[0];
                    let each_criteria_input;
                    for (each_criteria_input of criteria_div.getElementsByClassName("filter_criteria")) {
                        temp_filter["criteria"].push(each_criteria_input.value);
                    }
                    result["filter"]["filters"].push(temp_filter);
                }
            } else if (submode_name.startsWith("Copy")) {
                // copy
                console.log("copy");
                var col2col_div = each_submode.getElementsByClassName("col2cols")[0];
                var fix2col_div = each_submode.getElementsByClassName("fix2cols")[0];
                result["copy"] = {};
                result["copy"]["col2col"] = {};
                result["copy"]["fix2col"] = {};
                let each_copy;
                for (each_copy of col2col_div.getElementsByClassName("copy_relation_div")) {
                    if (each_copy.getElementsByClassName("copy_in1")[0].value === "") {
                        alert("Oops, seems that the target of the Copy part is empty!");
                        return null;
                    }
                    if (each_copy.getElementsByClassName("copy_in3")[0].value === "") {
                        alert("Oops, seems that the source of the Copy part is empty!");
                        return null;
                    }
                    var target_copy = each_copy.getElementsByClassName("copy_in1")[0].value.toUpperCase();
                    var source_copy = each_copy.getElementsByClassName("copy_in3")[0].value.toUpperCase();
                    result["copy"]["col2col"][target_copy] = source_copy;
                }
                for (each_copy of fix2col_div.getElementsByClassName("copy_relation_div")) {
                    if (each_copy.getElementsByClassName("copy_in1")[0].value === "") {
                        alert("Oops, seems that the target of the Copy part is empty!");
                        return null;
                    }
                    if (each_copy.getElementsByClassName("copy_in3")[0].value === "") {
                        alert("Oops, seems that the source of the Copy part is empty!");
                        return null;
                    }
                    var target_copy = each_copy.getElementsByClassName("copy_in1")[0].value.toUpperCase();
                    var source_copy = each_copy.getElementsByClassName("copy_in3")[0].value;
                    result["copy"]["fix2col"][target_copy] = source_copy;
                }
            } else if (submode_name.startsWith("Option")) {
                // options
                console.log("option");
                var append_true_checkbox = each_submode.getElementsByClassName("checkbox_append_true")[0];
                var append_false_checkbox = each_submode.getElementsByClassName("checkbox_append_false")[0];
                var screenupdating_true_checkbox = each_submode.getElementsByClassName("checkbox_screenupdating_true")[0];
                var screenupdating_false_checkbox = each_submode.getElementsByClassName("checkbox_screenupdating_false")[0];
                if (append_true_checkbox.checked && append_false_checkbox.checked) {
                    alert("Oops, seems that both TRUE and FALSE of the append in Options part are chosen.");
                    return null;
                }
                if (screenupdating_true_checkbox.checked && screenupdating_false_checkbox.checked) {
                    alert("Oops, seems that both TRUE and FALSE of the screen updating in Options part are chosen.");
                    return null;
                }
                if ( (! append_true_checkbox.checked) && (! append_false_checkbox.checked)) {
                    alert("Oops, seems that both TRUE and FALSE of the append in Options part are not chosen.");
                    return null;
                }
                if ( (! screenupdating_true_checkbox.checked) && (! screenupdating_false_checkbox.checked)) {
                    alert("Oops, seems that both TRUE and FALSE of the screen updating in Options part are not chosen.");
                    return null;
                }
                result["options"] = {};
                if (append_true_checkbox.checked) {
                    result["options"]["append"] = "true";
                } else {
                    result["options"]["append"] = "false";
                }
                if (screenupdating_true_checkbox.checked) {
                    result["options"]["disable_screen_updating"] = "true";
                } else {
                    result["options"]["disable_screen_updating"] = "false";
                }
            }
        }
        final_result["job"].push(result);
    }
    console.log(final_result);
    console.log("EnEnEn");
    console.log(JSON.stringify(final_result));
    // TODO: add POST here
    const xmlHttpRequest = new XMLHttpRequest();
    // testing for GET Request
    xmlHttpRequest.open("POST", "http://localhost:5000/", false);
    xmlHttpRequest.setRequestHeader("Content-type", "application/json");
    xmlHttpRequest.onload = () => {
        if (xmlHttpRequest.readyState === 4) {
            switch (xmlHttpRequest.status) {
                case 200:
                    console.log(xmlHttpRequest.responseText);
                    var body = document.getElementsByTagName("body")[0];
                    var code_block = document.createElement("code");
                    body.replaceChild(code_block, document.getElementById("conditions"));
                    code_block.innerHTML = xmlHttpRequest.responseText;
                    var new_button = document.createElement("button");
                    new_button.setAttribute("id", "download_button");
                    new_button.innerText = "Download Script";
                    document.getElementById("header").appendChild(new_button);
                    new_button.addEventListener("click", () => {
                       download("VBA_Script_Generated.txt", code_block.innerText);
                       alert("Downloading finished.");
                       location.reload();
                    });
                    break;
                default:
                    console.log(xmlHttpRequest.status);
                    alert("Oops, something wrong");
            }
        }
    };

    xmlHttpRequest.send(JSON.stringify(final_result));
    return final_result;
}

function download(filename, text) {
    var element = document.createElement("a");
    element.setAttribute("href", 'data:text/plain;charset=utf-8, ' + encodeURIComponent(text));
    element.setAttribute("download", filename);
    element.style.display = 'none';
    document.body.appendChild(element);
    element.click();
    document.body.removeChild(element);
}

var conditions = document.getElementById("conditions");
conditions.setAttribute("number", 0);
var add_button = document.getElementById("add_condition");
var submit_button = document.getElementById("submit");

add_button.addEventListener("click", () => {
    addCondition();
});

submit_button.addEventListener("click", () => {
    console.log("Submit");
    collect();
});
