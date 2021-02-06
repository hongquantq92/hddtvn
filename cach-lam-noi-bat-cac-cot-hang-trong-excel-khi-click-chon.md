Cách làm nổi bật các cột/hàng trong Excel khi click chọn
========================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-lam-noi-bat-cac-cot-hang-trong-excel-khi-click-chon/)Để làm rõ hơn nội dung, hoặc đánh dấu số liệu cần chú ý trong bảng dữ liệu Excel, chúng ta có tùy chọn tô màu xen kẽ các dòng Excel hoặc tô màu xen kẽ ô cho Excel, tùy theo nội dung cần chú ý ở vị trí nào. Khi đó màu sắc hiển thị xen …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Để làm rõ hơn nội dung, hoặc đánh dấu số liệu cần chú ý trong bảng dữ liệu Excel, chúng ta có tùy chọn tô màu xen kẽ các dòng Excel hoặc tô màu xen kẽ ô cho Excel, tùy theo nội dung cần chú ý ở vị trí nào. Khi đó màu sắc hiển thị xen kẽ cho tất cả đối tượng dữ liệu trong bảng. Trong trường hợp bạn không muốn tô màu xen kẽ mà chỉ muốn có màu khi click chuột trực tiếp thì có thể sử dụng code VBA. Với cách này bảng dữ liệu sẽ được nổi bật các giá trị theo hàng hoặc cột khi đối chiếu, bằng cách click chuột khi cần sử dụng mà thôi. Mời các bạn theo dõi bài viết sau để biết cách sử dụng VBA làm nổi bật cột và dòng trong Excel nhé.**


### 1. Làm nổi bật cột Excel


**Bước 1:** Để làm nổi bật dòng bằng cách tô màu dòng đó khi bạn nhấn chuột vào một ô bất kỳ, đầu tiên các bạn cần nhấn tổ hợp phím tắt **Alt + F11** để mở cửa sổ **Microsoft Visual Basic for Applications**.


**Bước 2:** Lúc này, cửa sổ VBA hiện ra. Các bạn nhấn đúp chuột trái vào sheet mà bạn muốn thao tác ở danh sách sheet bên trái. Sau đó sao chép đoạn code dưới đây vào.


Private Sub Worksheet\_SelectionChange(ByVal Target As Range)


On Error Resume Next


Cells.Interior.ColorIndex = 0


ActiveCell.EntireColumn.Interior.ColorIndex = 8


Application.CutCopyMode = True


End Sub


![](https://hddtvn.com/wp-content/uploads/2021/01/H0nQiQa.png)


**Bước 3:** Sau đó các bạn nhấn **Alt + Q** để tắt cửa sổ VBA đi. Sau đó quay lại trang tính Excel thì chỉ cần các bạn nhấn chuột vào một ô bất kỳ thì cả cột chứa ô đó sẽ được highlight lên.


![](https://hddtvn.com/wp-content/uploads/2021/01/wB8HWZf.png)


### 2. Làm nổi bật dòng Excel


**Bước 1:** Để làm nổi bật dòng bằng cách tô màu dòng đó khi bạn nhấn chuột vào một ô bất kỳ, đầu tiên các bạn cần nhấn tổ hợp phím tắt **Alt + F11** để mở cửa sổ **Microsoft Visual Basic for Applications**.


**Bước 2:** Lúc này, cửa sổ VBA hiện ra. Các bạn nhấn đúp chuột trái vào sheet mà bạn muốn thao tác ở danh sách sheet bên trái. Sau đó sao chép đoạn code dưới đây vào.


Private Sub Worksheet\_SelectionChange(ByVal Target As Range)


On Error Resume Next


Cells.Interior.ColorIndex = 0


ActiveCell.EntireRow.Interior.ColorIndex = 8


Application.CutCopyMode = True


End Sub


![](https://hddtvn.com/wp-content/uploads/2021/01/0agvcs9.png)


**Bước 3:** Sau đó các bạn nhấn **Alt + Q** để tắt cửa sổ VBA đi. Sau đó quay lại trang tính Excel thì chỉ cần các bạn nhấn chuột vào một ô bất kỳ thì cả dòng chứa ô đó sẽ được highlight lên.


![](https://hddtvn.com/wp-content/uploads/2021/01/A5bYv15.png)


### 3. Làm nổi bật cả cột và dòng Excel


**Bước 1:** Để làm nổi bật cả cột và dòng bằng cách tô màu dòng đó khi bạn nhấn chuột vào một ô bất kỳ, đầu tiên các bạn cần nhấn tổ hợp phím tắt **Alt + F11** để mở cửa sổ **Microsoft Visual Basic for Applications**.


**Bước 2:** Lúc này, cửa sổ VBA hiện ra. Các bạn nhấn đúp chuột trái vào sheet mà bạn muốn thao tác ở danh sách sheet bên trái. Sau đó sao chép đoạn code dưới đây vào.


Private Sub Worksheet\_SelectionChange(ByVal Target As Range)


On Error Resume Next


Cells.Interior.ColorIndex = 0


ActiveCell.EntireRow.Interior.ColorIndex = 8


ActiveCell.EntireColumn.Interior.ColorIndex = 8


Application.CutCopyMode = True


End Sub


![](https://hddtvn.com/wp-content/uploads/2021/01/luEj3hT.png)


**Bước 3:** Sau đó các bạn nhấn **Alt + Q** để tắt cửa sổ VBA đi. Sau đó quay lại trang tính Excel thì chỉ cần các bạn nhấn chuột vào một ô bất kỳ thì cả cột và dòng chứa ô đó sẽ được highlight lên.


![](https://hddtvn.com/wp-content/uploads/2021/01/ZZq8r4d.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng VBA để làm nổi bật cột và dòng trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreĐể làm rõ hơn nội dung, hoặc đánh dấu số liệu cần chú ý trong…

