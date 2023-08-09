<h1 align="center">
📄<br Export RVTools Script - Modified Version
</h1>

## 📚 Export RVTools Script - Modified Version

> RVTools is a Windows .NET (4.6.2 or higher) application which uses VMware vSphere Management SDK 8.0 and CIS REST API to display information about your virtual environments. Interacting with VirtualCenter 5.x, ESX Server 5.x, VirtualCenter 6.x, ESX Server 6.x, VirtualCenter 7.0, ESX server 7.0, VirtualCenter 8.0 and ESX server 8.0.  RVTools is able to list information about VMs, CPU, Memory, Disks, Partitions, Network, CD drives, USB devices. Snapshots, VMware tools, vCenter server,Resource pools, Clusters, ESX hosts, HBAs, Nics, Switches, Ports, Distributed Switches, Distributed Ports, Service consoles, VM Kernels, Datastores, multipath info, license info and health checks.

The information can be exported to csv and xlsx file(s). With a xlsx merge utility it’s possible to merge muliple vCenter xlsx reports to a single xlsx report.

When you install RVTOOLs, in default directory (C:\Program Files (x86)\Robware\RVTools) you have a script called "RVToolsBatchMultipleVCs.ps1" I modified this to include more actions. 

- Link for RVTOOLS - [Click Here To Download](https://www.robware.net/rvtools/)
- Link for RVTOOLS Version Info - [Click Here to Read](https://www.robware.net/rvtools/version/)

## Pre-Requirements

> Powershell version 5.1 or above

> Powercli version 10 or above

> RVTOOLS Version 4.4.1 (February 11, 2023) or above

## O que este script faz?

1. Conecta no(s) vCenter(s) especificado(s)
2. Cria pastas de acordo com ano, mês e dia para armazenar o export dos RVTOOLs (formato pt-br or en-us)
3. Se necessário, faz o merge dos arquivos
4. Se necessário, envia e-mail com os RVTOOLs gerados
5. Conecta com diferentes credenciais de acordo com algum parâmetro especificado

<div align="center">
  <br/>
  <br/>
  <br/>
    <div>
      <h1>Open Source</h1>
      <sub>Copyright © 2023 - <a href="https://github.com/julianoabr">julianoabr</sub></a>
    </div>
    <br/>
    📚
</div>



