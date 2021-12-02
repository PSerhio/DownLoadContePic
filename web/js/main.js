
async function do_report () {
    // получаем название файла
    let file1 = document.getElementById("formFile1").value.split('\\').pop();

    console.log(file1)
    await eel.do(file1)();
}

document.getElementById("btn-run").onclick = do_report;

// document.getElementById("btn-run").addEventListener("click", ()=>{do_report()}, false);