<div align="center">

## Get the Computername of a server in a Web Farm


</div>

### Description

We recently run into a problem working with a web farm and replication. We where getting a error that the tech people were say it was code and us the code people was saying it was Server related. Well, First we know it only happens 3 out of 10, So it must be happening only on one or two servers. If it was code It would happen on all servers, right???

Well, Lets find out which servers get the error, but how? ServerVariables only returns BROWSER header information and we need machine specific data. A Clustered farm generally uses the same IP info on a load balancer. so how do we get machine specific data. Well, use WSH of course. Three Lines. Hope you enjoy it.

This will not work on XP, because the disabled WSH by default. you can turn it on though. check MS Knowledge bases for instructions.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[John Tolar](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-tolar.md)
**Level**          |Advanced
**User Rating**    |3.8 (19 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Debugging and Error Handling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/debugging-and-error-handling__4-6.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/john-tolar-get-the-computername-of-a-server-in-a-web-farm__4-7837/archive/master.zip)





### Source Code

```
Set objWSHNet = CreateObject("WScript.Network")
Response.write objWSHNet.ComputerName & "<BR>"
Set objWSHSNet = Nothing
```

