### ZebulonVSTO (PowerPoint Add-in)

### I. IDE Settings

1. Visual Studio 2022

    - Community v17.5.2

    - Office/SharePoint 개발
    
    ```
    (포함됨) Visual Studio용 Office 개발자 도구
    (포함됨) .Net Framework 4.7.2 개발 도구
    (포함됨) Developer Analytics Tools
    (선택 사항) VSTO(Visual Studio Tools for Office)
    (선택 사항) 웹 배포:
    (선택 사항) .Net Framewrok 4.8 개발 도구
    ```

2. PowerPoint 추가기능 VSTO Project

    - PowerPorint 2013+

    - NuGet Packages

    ```
    Microsoft.Bcl.AsyncInterfaces
    System.Buffers
    System.Memory
    System.Numerics.Vectors
    System.Runtime.CompilerServices.Unsafe
    System.Text.Encodings.Web
    System.Text.Json
    System.Threading.Tasks.Extensions
    System.ValueTuple
    ```


### II. Features

1. Sync (UDP Sender/Receiver)

    - `alert [text]` : Open an alert dialog with the text message.

    - `select [n]` : Select the n-th slide in the editor window.

    - `showslide [n]` : Start the slide show and move to the n-th slide.

    - `hideslide` : Finish the slide show.