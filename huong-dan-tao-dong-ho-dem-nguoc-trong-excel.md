Hướng dẫn tạo đồng hồ đếm ngược trong Excel
===========================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-tao-dong-ho-dem-nguoc-trong-excel/)Tạo đồng hồ đếm ngược trong Excel đang là chủ đề được rất nhiều người quan tâm. Bạn sẽ cần tới đồng hồ đếm ngược khi phải “chạy tiến độ” hoàn thành một công việc hoặc dự án nào đó. Hãy theo dõi bài viết sau để biết cách tạo đồng hồ đếm ngược trong …
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Tạo đồng hồ đếm ngược trong Excel đang là chủ đề được rất nhiều người quan tâm. Bạn sẽ cần tới đồng hồ đếm ngược khi phải “chạy tiến độ” hoàn thành một công việc hoặc dự án nào đó. Hãy theo dõi bài viết sau để biết cách tạo đồng hồ đếm ngược trong Excel nhé.**


**Bước 1:** Để tạo đồng hồ đếm ngược, đầu tiên các bạn cần mở file Excel cần tạo đồng hồ lên. Sau đó các bạn nhấn chuột phải vào ô tạo đồng hồ đếm ngược. Thanh cuộn hiện ra các bạn chọn mục **Format Cells**.


![Hướng dẫn tạo đồng hồ đếm ngược trong Excel](https://hddtvn.com/wp-content/uploads/2021/01/bMGJpZD.png "Hướng dẫn tạo đồng hồ đếm ngược trong Excel")


**Bước 2:** Lúc này, hộp thoại **Format Cells** hiện ra. Các bạn chọn thẻ **Number**. Sau đó tại mục **Category** các bạn chọn **Time**. Tiếp theo tại mục **Locale** các bạn chọn **Vietnamese**. Sau đó chọn định dạng của giờ muốn tạo tại mục **Type**. Cuối cùng nhấn **OK** để hoàn tất.


![](https://hddtvn.com/wp-content/uploads/2021/01/ip778st.png)


**Bước 3:** Quay lại trang tính Excel, các bạn nhập thời gian muốn đếm ngược vào cũng như chỉnh sửa format của ô sao cho nổi bật cũng như tạo chú thích cho ô đó.


![](https://hddtvn.com/wp-content/uploads/2021/01/C5eHLgi.png)


**Bước 4:** Bây giờ các bạn chọn thẻ **Developer** trên thanh công cụ. Sau đó chọn mục **Visual Basic**. Hoặc các bạn có thể sử dụng tổ hợp phím tắt **Alt + F11** để mở cửa sổ VBA.


![](https://hddtvn.com/wp-content/uploads/2021/01/jvHdsEv.png)


**Bước 5:** Lúc này, cửa sổ **Microsoft Visual Basic for Applications** hiện ra. Các bạn chọn thẻ **Insert** trên thanh công cụ. Thanh cuộn hiện ra các bạn chọn mục **Module**.


![](https://hddtvn.com/wp-content/uploads/2021/01/lo58Yv2.png)


**Bước 6:** Lúc này hộp thoại Module hiện ra. Các bạn sao chép mã code dưới đây vào đó. Lưu ý D3 sẽ thay bằng vị trí ô mà bạn tạo đồng hồ.


Dim gCount As Date  

'Updateby20140925  

Sub Timer()  

gCount = Now + TimeValue("00:00:01")  

Application.OnTime gCount, "ResetTime"  

End Sub  

Sub ResetTime()  

Dim xRng As Range  

Set xRng = Application.ActiveSheet.Range("D3")  

xRng.Value = xRng.Value - TimeSerial(0, 0, 1)  

If xRng.Value <= 0 Then  

MsgBox "Countdown complete."  

Exit Sub  

End If  

Call Timer  

End Sub


#### **Bước 7:** Sau khi đã sao chép mã code xong.


Các bạn nhấn vào biểu tượng của **Run** trên thanh công cụ để mã code.


![](https://hddtvn.com/wp-content/uploads/2021/01/JB1HWg7.png)


Chỉ cần như vậy là thời gian tại ô D3 sẽ tự động đếm ngược. Các bạn có thể thay đổi lại thời gian đếm ngược tùy ý và đồng hồ sẽ không bị dừng lại vẫn tự đếm ngược đến khi hết giờ.


![](https://hddtvn.com/wp-content/uploads/2021/01/9Hl1vnT.png)


Sau khi thời gian đếm ngược kết thúc thì sẽ có thông báo **Countdown complete** hiện ra như hình dưới.


![Hướng dẫn tạo đồng hồ đếm ngược trong Excel](https://hddtvn.com/wp-content/uploads/2021/01/DEYwFAK.png "Hướng dẫn tạo đồng hồ đếm ngược trong Excel")


*Như vậy, bài viết trên đã hướng dẫn các bạn cách tạo đồng hồ đếm ngược trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreTạo đồng hồ đếm ngược trong Excel đang là chủ đề được rất nhiều người…

