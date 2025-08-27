# SAP ECC Automation

Sometimes it is necessary to automate SAP in environments where SAP GUI Scripting is locked down/disabled. This library can be used to automate SAP GUI in these conditions.

## Typical usage

I've used these libraries mostly to perform extraction tasks:

* Extract SAP Notification long text (Service and PM)
* Extract SAP Order long text
* Extract tables from IH06
* Extract information from address form from functional locations in IH06

I would generally advise against using this library for any tasks which update financial information. Do so at your own risk.

## Note:

In order to use the library you can either Create a VBA instance of SAP ECC from existing SAP window, or a new process. To create a new process the current library expects you to launch SAP ECC from a Web URL. You may want to use a different mechanism. Ideally this library would accommodate for that but without experience what that looks like I do not have the ability to build such a routine. E.G. I know some people need to use user/password.

## Library dependencies

### Base:

* stdICallable
* stdAcc
* stdWindow
* stdClipboard
* stdLambda

### With Async:

Additional modules are required for asynchronous processing. Asynchronous processing allows you to automate multiple SAP windows at once. Care must be had while doing these kind of operations though as SAP is glitchy and often actions aren't registered.

* stdFiber
* stdProcess
* stdCallback

## High Level Process

Typical process flow for loading and exporting several variants from IH06:

```mermaid
flowchart TD
    B[Create IH06 Instance sapSAPECCIH06.CreateSync]
    B --> B1[Ensure GuiXT Enabled]
    B1 --> C[Initialise SAP ECC Session Transaction = IH06]

    C --> D[Loop over Variants e.g. 3 times]
    D --> E[Load Variant using sapSAPECCIH06.loadVariqant]
    E --> F[Execute Search and await results]
    F --> G[Export Results → CSV using sapECCIH06.exportAsSpreadsheet]
    G --> D

    D -->|All Variants Processed| H[Return to Home / Reset GuiXT]
    H --> I[Complete]
```

## Project Structure

```mermaid
flowchart LR
    subgraph BaseLibraries[stdVBA Utilities]
        SA[stdAcc]
        SW[stdWindow]
        SCB[stdClipboard]
        SI[stdICallable]
        SL[stdLambda]
        SR[stdReg]
    end

    subgraph SAPCore[SAP ECC Core]
        SE[sapSAPECC]
        
    end

    subgraph IH06[IH06 Automation]
        IH[sapSAPECCIH06]
    end

    %% Dependencies
    IH --> SE
    IH --> SA
    IH --> SW
    IH --> SCB
    IH --> SL
    IH --> SR

    SE --> SA
    SE --> SW
    SE --> SCB
    SE --> SI
```
