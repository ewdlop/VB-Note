# **VBA 中的機器學習**

VBA (Visual Basic for Applications) 是一種程式語言，主要用於自動化 Microsoft Office 應用程式中的任務，例如 Excel。VBA 對於需要處理大量資料集並需要頻繁產生報表的專業人士特別有用，例如財務會計師或行政助理。 1 透過自動執行重複性任務，VBA 可以顯著提高效率和生產力。雖然 VBA 並非專為機器學習設計，但它可以透過一些方法來執行基本的機器學習任務。 2

## **研究方法**

本文基於廣泛的研究，探索各種線上資源，包括與 VBA 和機器學習相關的技術部落格、線上課程和書籍。研究涉及搜尋如何在 VBA 中使用機器學習的資訊、識別相關函式庫和工具，以及分析程式碼範例和限制。

## **如何在 VBA 中使用機器學習**

在 VBA 中使用機器學習，主要有以下幾種方法：

* **使用 VBA 執行簡單的機器學習演算法：** VBA 可以用來執行一些簡單的機器學習演算法，例如線性迴歸、決策樹和 K-Nearest Neighbors 演算法等。 3 K-Nearest Neighbors 演算法是一種用於分類和迴歸的演算法，它透過計算資料點之間的距離來進行預測。例如，在 Excel 中，可以使用 K-Nearest Neighbors 演算法來根據歷史資料預測客戶的信用風險。這些演算法可以透過 VBA 程式碼來實現，並用於分析 Excel 中的資料。 4  
* **使用 VBA 與外部機器學習工具整合：** VBA 可以與其他程式語言和工具整合，例如 Python。 4 與需要昂貴費用和複雜實作的專用 RPA 軟體相比，VBA 提供了一種經濟高效的方式在 Excel 中自動執行任務。透過 VBA 呼叫 Python 腳本，可以利用 Python 強大的機器學習庫 (例如 scikit-learn) 來執行更需多資源的機器學習任務。 2 這樣可以讓使用者在 Excel 中處理資料，並使用 Python 進行機器學習模型的訓練和預測。例如，使用者可以透過 VBA 將 Excel 中的資料預處理後，傳遞給 Python 中的機器學習模型進行訓練，最後將預測結果匯回 Excel。  
* **使用 AI 工具輔助 VBA 程式碼開發：** AI 工具，例如 ChatGPT，可以協助使用者產生 VBA 程式碼，自動化工作流程。 5 AI 工具，例如 ChatGPT，可以大幅降低 VBA 的學習難度，方法是提供程式碼協助、範例和除錯支援，讓初學者更容易上手自動化。此外，AI 工具也可以提供程式碼範例和除錯建議，幫助使用者更快地學習 VBA。 6

值得注意的是，VBA 可以被視為通往更複雜自動化解決方案（如機器人流程自動化 (RPA)）的墊腳石。RPA 使用軟體機器人跨各種應用程式自動執行重複性任務，而 VBA 可用於在 Excel 環境中建置 RPA 工作流程。 7

在應用機器學習演算法之前，分析和準備資料至關重要。VBA 可以透過允許使用者執行資料分析任務（例如清理資料、將資料轉換為合適的格式以及產生視覺化以獲取洞察力）在這個過程中發揮重要作用。 2

## **AI 工具 for VBA and Machine Learning**

如前所述，AI 工具在簡化 VBA 學習和應用方面發揮著越來越重要的作用。以下是一些值得注意的 AI 工具及其優勢：

* **ChatGPT：** 作為一種大型語言模型，ChatGPT 可以根據使用者的需求產生 VBA 程式碼、提供程式碼範例，甚至協助除錯。這顯著降低了 VBA 的學習門檻，讓沒有程式設計經驗的使用者也能夠使用 VBA 自動執行任務。 6  
* **其他 AI 輔助工具：** 除了 ChatGPT 之外，還有其他 AI 工具可以協助 VBA 程式碼開發，例如可以根據使用者輸入的程式碼片段自動產生程式碼建議的工具，以及提供 VBA 教學和指南的 AI 學習資源。 4 這些工具可以幫助使用者更快地完成 VBA 程式碼編寫，同時降低錯誤的風險。

## **VBA 機器學習函式庫和工具**

VBA 缺乏專為矩陣運算、影像處理和類神經網路等複雜機器學習任務而設計的內建函式庫。 3 但是，一些書籍和線上課程提供了使用 VBA 進行機器學習的教學和範例程式碼。 8 例如，《AI 輔助學習 Excel VBA 最強入門》這本書，以循序漸進的方式引導讀者學習 VBA，並透過 AI 輔助學習的方式，讓學習更有效率。 8

此外，一些線上課程，例如台大資訊系統訓練班的「ChatGPT x VBA 自動化工作流程」課程，教導學員如何使用 ChatGPT 自動產生 VBA 程式碼，並開發自訂函數和自動排程等功能。 9

## **VBA 機器學習的限制**

雖然 VBA 可以用於執行基本的機器學習任務，但也存在一些限制：

* **執行效能：** VBA 的執行速度比其他程式語言 (例如 Python) 慢。 3 這對於需要處理大量資料或執行複雜演算法的機器學習任務來說，可能會是一個問題。  
* **資料處理能力：** VBA 對於不同類型資料的處理能力較弱，例如文字、圖片等。 3 這限制了 VBA 在機器學習方面的應用範圍。  
* **函式庫支援：** VBA 缺乏機器學習所需的函式庫支援，例如矩陣運算、圖片處理等。 3 這表示使用者需要自行開發這些函式，才能執行機器學習任務。  
* **名稱長度：** 每個宣告的程式設計元素名稱都有字元數上限。如果元素名稱是限定的，則此最大值適用於整個限定字串。 10  
* **行長度：** 原始程式碼的實體行中最多只能有 65,535 個字元。如果使用行接續字元，邏輯原始程式碼可能會更長。 10  
* **陣列維度：** 您可以為陣列宣告的維度數目設有上限。這會限制您可以使用多少索引來指定陣列元素。 10  
* **字串長度：** 您可以在單一字串中儲存的 Unicode 字元數目設有上限。 10

這些限制可能會影響 VBA 中機器學習任務的執行效率和複雜度。例如，名稱長度限制可能會使程式碼難以閱讀和維護，而字串長度限制可能會影響處理大量文字資料的能力。

## **VBA 機器學習程式碼範例**

以下是一些 VBA 機器學習程式碼範例：

* **自動化輸入文字或數值：** 可以透過程式碼，讓 Excel 自動輸入大量的文字或數值資料。這在機器學習中很有用，例如自動填寫缺失值或產生訓練資料。 1  
* **自動匯入檔案：** 可以使用 VBA 程式碼自動匯入外部檔案到 Excel 中。這對於機器學習任務來說非常重要，因為通常需要從外部來源匯入資料進行分析和模型訓練。 7  
* **自動建立工作表：** 可以使用 VBA 程式碼自動建立新的工作表。這在機器學習中很有用，例如將資料集分成訓練集和測試集，或儲存不同模型的預測結果。 7  
* **找尋最後一列：** 可以使用 VBA 程式碼找到工作表中最後一列的資料。這在處理動態資料集時很有用，例如需要根據最新的資料更新機器學習模型。 7  
* **複製、貼上資料：** 可以使用 VBA 程式碼複製和貼上資料。這在資料預處理和特徵工程中很有用，例如需要將資料轉換為不同的格式或建立新的特徵。 7  
* **自動輸入公式：** 可以使用 VBA 程式碼自動輸入公式，並自動填滿到最後一列。這在特徵工程中很有用，例如需要計算新的特徵或轉換現有特徵。 7

建立後，這些 VBA 巨集可以儲存在 Excel 活頁簿中，並在需要時重複使用，從而進一步提高生產力並節省時間。 7

## **Real-world Applications**

VBA 和機器學習的結合可以在各種實際場景中發揮作用。以下是一些例子：

* **客戶流失預測：** VBA 可以用於自動化資料預處理，例如清理資料、轉換資料格式和產生視覺化，以便在 Python 中建置客戶流失預測模型。  
* **銷售預測：** VBA 可以用於從 Excel 中提取銷售資料，並將其傳遞給 Python 中的機器學習模型進行銷售預測。預測結果可以匯回 Excel，以便使用者進行進一步分析和決策。  
* **庫存管理：** VBA 可以用於自動化庫存資料的更新和分析，並使用機器學習模型預測未來的庫存需求。  
* **風險評估：** VBA 可以用於收集和分析風險相關資料，並使用機器學習模型評估不同投資或專案的風險。

## **如何在 VBA 中評估機器學習模型**

評估機器學習模型的效能是機器學習流程中重要的一環。 11 在 VBA 中，可以使用一些基本的評估指標來評估模型的效能，例如：

* **準確率 (Accuracy)：** 計算模型預測正確的比例。 12  
* **F1 Score：** 綜合考量精確率 (Precision) 和召回率 (Recall) 的指標。 12  
* **ROC 曲線和 AUC：** ROC 曲線是用來評估二元分類模型的效能，AUC 是 ROC 曲線下的面積，AUC 越大，模型的效能越好。 12  
* **混淆矩陣 (Confusion Matrix)：** 用於顯示模型預測結果和實際結果的表格，可以幫助使用者了解模型的預測錯誤類型。 13

## **總結**

VBA 可以用於執行基本的機器學習任務，但它並非專為機器學習設計，因此存在一些限制。透過 VBA 與外部機器學習工具整合，例如 Python 的 scikit-learn 函式庫，或使用 AI 工具輔助 VBA 程式碼開發，可以克服這些限制，並執行更複雜的機器學習任務。 2 儘管 VBA 有其局限性，但它仍然是自動化 Excel 任務和執行基本機器學習任務的寶貴工具，特別是與 AI 工具和外部函式庫整合時。本文的主要結論包括：

* VBA 提供了一種經濟高效的方式在 Excel 中自動執行任務，使其成為自動化的良好起點。  
* 資料分析在機器學習過程中至關重要，VBA 可以協助執行資料清理、轉換和視覺化等任務。  
* AI 工具可以簡化 VBA 學習，並協助產生程式碼和除錯。  
* 將 VBA 與外部機器學習函式庫整合可以增強其功能，並允許執行更複雜的機器學習任務。

## **額外資訊**

| 資源 | 描述 |
| :---- | :---- |
| AI 輔助學習 Excel VBA 最強入門 | 一本創新的 VBA 學習指南，透過 AI 輔助學習的方式，讓學習更有效率。 |
| ChatGPT x VBA 自動化工作流程 | 台大資訊系統訓練班的線上課程，教導學員如何使用 ChatGPT 自動產生 VBA 程式碼。 |
| Excel VBA 範例字典 | 提供大量的 VBA 程式碼範例，幫助使用者學習 VBA。 |
| Microsoft Learn | Microsoft 官方網站，提供 VBA 的教學文件和程式碼範例。 |
| Excel Home 技術論壇 | Excel 使用者交流平台，可以找到 VBA 相關的教學和問題解答。 |

#### **Works cited**

1\. VBA是什麼？寫給入門者的Excel Visual Basic教學文章 \- 巨匠電腦, accessed February 24, 2025, [https://www.pcschool.com.tw/blog/it-skill/excel-vba-tutorial](https://www.pcschool.com.tw/blog/it-skill/excel-vba-tutorial)  
2\. VBA语言的人工智能原创 \- CSDN博客, accessed February 24, 2025, [https://blog.csdn.net/2501\_90485535/article/details/145447649](https://blog.csdn.net/2501_90485535/article/details/145447649)  
3\. 能用Excel做機器學習嗎 \- Edony A.I., accessed February 24, 2025, [https://edonyai.home.blog/2019/11/12/%E8%83%BD%E7%94%A8excel%E5%81%9A%E6%A9%9F%E5%99%A8%E5%AD%B8%E7%BF%92%E5%97%8E/](https://edonyai.home.blog/2019/11/12/%E8%83%BD%E7%94%A8excel%E5%81%9A%E6%A9%9F%E5%99%A8%E5%AD%B8%E7%BF%92%E5%97%8E/)  
4\. 【VBA入門指南】學懂VBA解鎖Excel的無限潛能| 提升生產力系列 \- 毅思會計, accessed February 24, 2025, [https://acaccountinghk.com/startup/vba/](https://acaccountinghk.com/startup/vba/)  
5\. AI輔助學習Excel VBA最強入門邁向辦公室自動化之路王者歸來下冊(二版) \- 博客來, accessed February 24, 2025, [https://www.books.com.tw/products/0010985865](https://www.books.com.tw/products/0010985865)  
6\. 有了AI，学习VBA的难度直降90%，弯道超车的机会又来了 \- ExcelHome, accessed February 24, 2025, [https://www.excelhome.net/4922.html](https://www.excelhome.net/4922.html)  
7\. 【Excel 宗師】用VBA 寫支機器人 . 實現資料處理 \- Medium, accessed February 24, 2025, [https://medium.com/@chunlin-damien-yu/excel-%E5%A4%A7%E5%B8%AB%E7%8F%AD-%E6%88%91%E7%94%A8-vba-%E5%AF%AB%E4%BA%86%E4%B8%80%E6%94%AF%E6%A9%9F%E5%99%A8%E4%BA%BA-%E5%AF%A6%E7%8F%BE%E8%B3%87%E6%96%99%E8%99%95%E7%90%86%E8%A7%A3%E6%94%BE%E9%9B%99%E6%89%8B%E7%9A%84%E6%9C%80%E5%BE%8C%E4%B8%80%E5%93%A9%E8%B7%AF-399ef5cdaad9](https://medium.com/@chunlin-damien-yu/excel-%E5%A4%A7%E5%B8%AB%E7%8F%AD-%E6%88%91%E7%94%A8-vba-%E5%AF%AB%E4%BA%86%E4%B8%80%E6%94%AF%E6%A9%9F%E5%99%A8%E4%BA%BA-%E5%AF%A6%E7%8F%BE%E8%B3%87%E6%96%99%E8%99%95%E7%90%86%E8%A7%A3%E6%94%BE%E9%9B%99%E6%89%8B%E7%9A%84%E6%9C%80%E5%BE%8C%E4%B8%80%E5%93%A9%E8%B7%AF-399ef5cdaad9)  
8\. AI輔助學習Excel VBA最強入門邁向辦公室自動化之路王者歸來上冊（二版） \- 博客來, accessed February 24, 2025, [https://www.books.com.tw/products/0010985863](https://www.books.com.tw/products/0010985863)  
9\. ChatGPT x VBA 自動化工作流程 \- 臺灣大學資訊系統訓練班, accessed February 24, 2025, [https://train.csie.ntu.edu.tw/train/course.php?id=5123](https://train.csie.ntu.edu.tw/train/course.php?id=5123)  
10\. Visual Basic 的限制 \- Microsoft Learn, accessed February 24, 2025, [https://learn.microsoft.com/zh-tw/dotnet/visual-basic/programming-guide/program-structure/limitations](https://learn.microsoft.com/zh-tw/dotnet/visual-basic/programming-guide/program-structure/limitations)  
11\. 使用機器學習解決問題的五步驟: 模型評估 \- DataSci Ocean, accessed February 24, 2025, [https://datasciocean.tech/machine-learning-basic-concept/machine-learning-model-evaluate/](https://datasciocean.tech/machine-learning-basic-concept/machine-learning-model-evaluate/)  
12\. Python機器學習-分類模型的5個評估指標 \- Medium, accessed February 24, 2025, [https://medium.com/@imirene/python%E6%A9%9F%E5%99%A8%E5%AD%B8%E7%BF%92-%E5%88%86%E9%A1%9E%E6%A8%A1%E5%9E%8B%E7%9A%845%E5%80%8B%E8%A9%95%E4%BC%B0%E6%8C%87%E6%A8%99-3260f116ce47](https://medium.com/@imirene/python%E6%A9%9F%E5%99%A8%E5%AD%B8%E7%BF%92-%E5%88%86%E9%A1%9E%E6%A8%A1%E5%9E%8B%E7%9A%845%E5%80%8B%E8%A9%95%E4%BC%B0%E6%8C%87%E6%A8%99-3260f116ce47)  
13\. 機器學習-常見的評估指標 \- MaDi's Blog, accessed February 24, 2025, [https://dysonma.github.io/2020/12/05/%E6%A9%9F%E5%99%A8%E5%AD%B8%E7%BF%92-%E5%B8%B8%E8%A6%8B%E7%9A%84%E8%A9%95%E4%BC%B0%E6%8C%87%E6%A8%99/](https://dysonma.github.io/2020/12/05/%E6%A9%9F%E5%99%A8%E5%AD%B8%E7%BF%92-%E5%B8%B8%E8%A6%8B%E7%9A%84%E8%A9%95%E4%BC%B0%E6%8C%87%E6%A8%99/)
