$data = get-date
get-MessageTrackingLog -Server elefante2 -EventId RECEIVE -ResultSize Unlimited -Start $data.AddMinutes(-5) > listaR.txt
get-MessageTrackingLog -Server elefante2 -EventId SEND -ResultSize Unlimited -Start $data.AddMinutes(-5) > listaS.txt
get-content listaR.txt | Measure-Object -Line > countR
get-content listaS.txt | Measure-Object -Line > countS
(get-content countR -TotalCount 4)[-1].substring(26,28) > qtdR.txt
(get-content countS -TotalCount 4)[-1].substring(26,28) > qtdS.txt
echo $data
