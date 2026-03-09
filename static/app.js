function parsePlan(){

    const md = document.getElementById("mdInput").value

    fetch("/parse",{
        method:"POST",
        headers:{
            "Content-Type":"application/json"
        },
        body:JSON.stringify({
            md:md
        })
    })
    .then(r=>r.json())
    .then(data=>{

        document.getElementById("result").innerText = data.result

    })

}

function evaluate(){

    const jsonText = document.getElementById("result").innerText

    fetch("/evaluate",{
        method:"POST",
        headers:{
            "Content-Type":"application/json"
        },
        body: jsonText
    })
    .then(r=>r.json())
    .then(data=>{
        console.log(data)

        document.getElementById("evaluation").innerText =
            JSON.stringify(data,null,2)
    })

}