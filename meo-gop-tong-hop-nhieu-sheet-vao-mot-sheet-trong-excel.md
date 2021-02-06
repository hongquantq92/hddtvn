Mẹo gộp/tổng hợp nhiều sheet vào một sheet trong Excel
======================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/meo-gop-tong-hop-nhieu-sheet-vao-mot-sheet-trong-excel/)Bạn cần gộp dữ liệu trong nhiều sheet vào chung một sheet? Bạn lo sợ copy/paste từng sheet sẽ rất tốn thời gian và dễ xảy ra sai sót cho file tổng hợp? Hãy theo dõi bài viết sau để biết cách gộp nhiều sheet vào một sheet trong Excel một cách nhanh chóng và …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Bạn cần gộp dữ liệu trong nhiều sheet vào chung một sheet? Bạn lo sợ copy/paste từng sheet sẽ rất tốn thời gian và dễ xảy ra sai sót cho file tổng hợp? Hãy theo dõi bài viết sau để biết cách gộp nhiều sheet vào một sheet trong Excel một cách nhanh chóng và đơn giản.


Ví dụ ta cần gộp 3 sheet A, B, C vào thành một sheet như hình dưới.


![](https://hddtvn.com/wp-content/uploads/2021/01/pO0AWM0.png)


Cách gộp các sheet nhanh nhất là sử dụng code VBA. Code VBA có thể giúp bạn gộp dữ liệu của tất cả sheet trong file Excel hiện tại thành một sheet mới. Điều kiện cần thiết là tất cả trang tính phải có cùng cấu trúc, cùng tiêu đề cột và thứ tự cột. Các bạn hãy làm theo các bước sau:


**Bước 1:** Đầu tiên các bạn nhấn tổ hợp phím tắt **Alt + F11** để mở cửa sổ của VBA.


![](https://hddtvn.com/wp-content/uploads/2021/01/u523PQ2.png)


**Bước 2:** Tiếp theo, các bạn chọn thẻ **Insert**. Thanh cuộn hiện ra các bạn chọn mục **Module**.


[![](https://hddtvn.com/wp-content/uploads/2021/01/BodR7tL.png)](https://hddtvn.com/wp-content/uploads/2021/01/BodR7tL.png)


**Bước 3:** Cửa sổ Module hiện ra, các bạn **copy** đoạn code dưới đây rồi **paste** vào cửa sổ.


Sub Combine()
 Dim J As Integer
 On Error Resume Next
 Sheets(1).Select
 Worksheets.Add
 Sheets(1).Name = "Combined"
 Sheets(2).Activate
 Range("A1").EntireRow.Select
 Selection.Copy Destination:=Sheets(1).Range("A1")
 For J = 2 To Sheets.Count
 Sheets(J).Activate
 Range("A1").Select
 Selection.CurrentRegion.Select
 Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1).Select
 Selection.Copy Destination:=Sheets(1).Range("A65536").End(xlUp)(2)
 Next
End Sub
**Bước 4:** Sau đó các bạn nhấn **F5** để chạy mã.


![](https://hddtvn.com/wp-content/uploads/2021/01/s14DvUM.png)


Chỉ cần như vậy là dữ liệu trong tất cả sheet trong file Excel đã được gộp vào một sheet tên là **Combined** được đặt trước các sheet khác.


![](https://hddtvn.com/wp-content/uploads/2021/01/HTJynpW.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách gộp nhiều sheet vào một sheet trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreBạn cần gộp dữ liệu trong nhiều sheet vào chung một sheet? Bạn lo sợ…

