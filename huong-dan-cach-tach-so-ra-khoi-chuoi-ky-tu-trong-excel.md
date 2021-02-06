Hướng dẫn cách tách số ra khỏi chuỗi ký tự trong Excel
======================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-cach-tach-so-ra-khoi-chuoi-ky-tu-trong-excel/)Bạn được giao một bảng dữ liệu chứa thông tin khách hàng và được yêu cầu chỉ lấy phần số (số điện thoại, ID khách hàng, số CMND, năm sinh….) nhằm trích xuất các dữ liệu này về file đặc thù để giao dịch hoặc liên hệ. Bạn sẽ làm thế nào để tách số …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Bạn được giao một bảng dữ liệu chứa thông tin khách hàng và được yêu cầu chỉ lấy phần số (số điện thoại, ID khách hàng, số CMND, năm sinh….) nhằm trích xuất các dữ liệu này về file đặc thù để giao dịch hoặc liên hệ. Bạn sẽ làm thế nào để tách số trong một bảng dữ liệu lớn? Hãy theo dõi bài viết sau để biết cách thực hiện nhé.**


Ví dụ ta có bảng dữ liệu như hình dưới. Yêu cần cần tách số chứng minh thư của từng người ra. Do dữ liệu không đồng nhất nên ta sẽ không thể sử dụng các hàm tách ký tự cơ bản như RIGHT, MID, LEFT được. Các bạn hãy làm theo các bước sau để tách nhé.


![](https://hddtvn.com/wp-content/uploads/2021/01/2SoLbvA.png)


**Bước 1:** Đầu tiên các bạn cần sử dụng tổ hợp phím tắt **Alt + F11** để mở cửa sổ Microsoft Visual Basic for Applications. Sau đó các bạn chọn thẻ **Insert** trên thanh công cụ. Thanh cuộn hiện ra các bạn chọn mục **Module**.


![](https://hddtvn.com/wp-content/uploads/2021/01/fQ6qHOV.png)


**Bước 2:** Tiếp theo, các bạn bôi đen toàn bộ đoạn code sau rồi nhấn chuột phải và chọn Copy hoặc sử dụng tổ hợp phím tắt Ctrl + C để sao chép.


Function ExtractNumber(rCell As Range)  

Dim lCount As Long  

Dim sText As String  

Dim lNum As String  

sText = rCell  

For lCount = Len(sText) To 1 Step -1  

If IsNumeric(Mid(sText, lCount, 1)) Then  

lNum = Mid(sText, lCount, 1) & lNum  

End If  

Next lCount  

ExtractNumber = CLng(lNum)  

End Function


**Bước 3:** Quay lại cửa sổ Module vừa được mở ra, các bạn nhấn chuột phải rồi chọn **Paste** hoặc sử dụng tổ hợp phím tắt **Ctrl + V** để dán đoạn code được sao chép ở trên vào đây.


![](https://hddtvn.com/wp-content/uploads/2021/01/VkgFJbq.png)


**Bước 4:** Chỉ cần như vậy là hàm **ExtractNumber** đã được thêm vào Excel của bạn. Bây giờ quay về trang tính các bạn sử dụng công thức sau:


**=ExtractNumber(B2)**


Sao chép công thức cho các đối tượng bên dưới ta sẽ thu được kết quả là số chứng minh thư nhân dân của từng người đã được tách ra một cách nhanh chóng.


![](https://hddtvn.com/wp-content/uploads/2021/01/HO44YER.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách tách số ra khỏi chuỗi ký tự trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreBạn được giao một bảng dữ liệu chứa thông tin khách hàng và được yêu…

