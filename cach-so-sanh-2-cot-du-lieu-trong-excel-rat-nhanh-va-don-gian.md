Cách so sánh 2 cột dữ liệu trong Excel: Rất nhanh và đơn giản
=============================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-so-sanh-2-cot-du-lieu-trong-excel-rat-nhanh-va-don-gian/)Khi muốn so sánh 2 cột dữ liệu thì đa phần các bạn sẽ so sánh thủ công bằng mắt thường. Nhưng cách làm này rất thời gian, dễ gây nhầm lẫn và sai sót nếu các cột của bạn có quá nhiều số liệu hoặc sử dụng các nội dung tương đồng nhau. Excel …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Khi muốn so sánh 2 cột dữ liệu thì đa phần các bạn sẽ so sánh thủ công bằng mắt thường. Nhưng cách làm này rất thời gian, dễ gây nhầm lẫn và sai sót nếu các cột của bạn có quá nhiều số liệu hoặc sử dụng các nội dung tương đồng nhau.** **Excel đã cung cấp một số hàm rất tiện dụng để phục vụ cho mục đích so sánh dữ liệu thông tin trên 2 cột, nhưng không phải ai cũng biết điều đó. Bài viết dưới đây sẽ hướng dẫn các bạn cách so sánh 2 cột dữ liệu trong Excel một cách dễ dàng, tiện lợi và hiệu quả nhất.**


![Cách so sánh 2 cột dữ liệu trong Excel dễ dàng, hiệu quả.](https://hddtvn.com/wp-content/uploads/2021/01/so-sanh-cot-du-lieu.png "Cách so sánh 2 cột dữ liệu trong Excel dễ dàng, hiệu quả.")


Cách sử dụng hàm COUNTIF trong Excel để so sánh 2 cột dữ liệu
-------------------------------------------------------------


Hàm **COUNTIF** là một trong các hàm thống kê. Được sử dụng để đếm số lượng ô đáp ứng một tiêu chí.


Cú pháp hàm: **COUNTIF( Vùng các bạn muốn đếm hay thống kê, Tiêu chí)**


Đối với từng phiên bản Excel mà chúng ta sử dụng dấu **;** hoặc dấu **,** trong công thức. Nếu các bạn thấy báo lỗi khi nhập hãy đổi lại dấu.


**Bước 1:** Các bạn mở file Excel chứa dữ liệu trong 2 cột cần so sánh.


![](https://hddtvn.com/wp-content/uploads/2021/01/77-1024x528-1.png)


**Bước 2:** Các bạn bôi đen vùng dữ liệu ở cột 1, trừ ô tiêu đề cột. Sau đó, click chuột vào **Name Box** để đổi tên. Lưu ý tên các bạn đặt là tên không dấu, viết liền không cách. Sau đó, các bạn gõ **Enter**.


![](https://hddtvn.com/wp-content/uploads/2021/01/78-1024x530-1.png)


#### **Bước 3:** Tương tự như vậy, các bạn cũng làm cho cột dữ liệu thứ 2 của mình.


![](https://hddtvn.com/wp-content/uploads/2021/01/79-1024x526-1.png)


**Bước 4**: Tại ô **D6**, các bạn nhập hàm. **COUNTIF(xuong2;C6)**. Ở công thức này được hiểu là vùng dữ liệu các bạn đang muốn đếm hay thống kê là **xuong2**, và tiêu chí so sánh là ô **C6**.


**Trong đó**, kết quả trả về **0** là không có dữ liệu nào ở cột Phân xưởng 2 trùng với dữ liệu ở cột Phân xưởng 1. Ví dụ, Áo phao siêu nhẹ chỉ có ở cột Phân xưởng 1, mà không có ở cột Phân xưởng 2. Nếu kết quả trả về là **1,2…** thì có nghĩa là số lần dữ liệu trùng nhau ở 2 cột Phân xưởng 1 và Phân xưởng 2. Ví dụ, Quần Jean xuất hiện 1 lần ở cả 2 cột dữ liệu, nên đã trả về kết quả là 1.


![](https://hddtvn.com/wp-content/uploads/2021/01/80-1024x528-1.png)


**Bước 5:** Tương tự như vậy, các bạn sẽ so sánh 2 cột dữ liệu cột Phân xưởng 1 và Phân xưởng 2 với tiêu chí là các ô tiếp theo **C7, C8, C9, C10**.


![](https://hddtvn.com/wp-content/uploads/2021/01/81-1024x528-1.png)


**Bước 6:** Trên phần **Name Box**, các bạn bấm vào mũi tên sổ xuống, và chọn xuong1.


![](https://hddtvn.com/wp-content/uploads/2021/01/82-1024x527-1.png)


**Bước 7:** Trên tab **Home**, các bạn kích chuột vào mũi tên sổ xuống ở **Conditional Formatting**. Sau đó, kích chọn **New Rule**.


![](https://hddtvn.com/wp-content/uploads/2021/01/83-1024x527-1.png)


**Bước 8:** Lúc này trên màn hình, hiển thị giao diện **New Formatting Rule**. Các bạn kích chuột chọn **Use a formula to determine which cells to format**. Tiếp theo, các bạn nhập công thức **=COUNTIF(xuong2;C6)=0** vào ô **Format values where this formula is true**. Sau đó, các bạn kích chuột vào **Format**.


![](https://hddtvn.com/wp-content/uploads/2021/01/88.png)


**Bước 9:** Màn hình hiển thị hộp thoại **Format Cells**. Các bạn vào tab **Fill**, sau đó chọn màu sắc mình muốn, rồi nhấn **OK**


![](https://hddtvn.com/wp-content/uploads/2021/01/85.png)


Các bạn kiểm tra lại các thông tin trên hộp thoại **New Formatting Rule**, sau đó nhấn **OK**


![](https://hddtvn.com/wp-content/uploads/2021/01/89.png)


#### Các bạn sẽ nhận được kết quả trả về như hình dưới. Các mã hàng hay dữ liệu xuất hiện ở cột Phân xưởng 1 mà không có ở cột Phân xưởng 2 sẽ được đổ màu xanh.


![](https://hddtvn.com/wp-content/uploads/2021/01/90-1024x545-1.png)


**Bước 10:** Các bạn làm tương tự với cột 2. Công thức trong **Format values where this formula is true** là **=COUNTIF(xuong1;F6)=0**.


![](https://hddtvn.com/wp-content/uploads/2021/01/91.png)


Kết quả, các bạn sẽ nhận được thêm dữ liệu xuất hiện ở cột Phân xưởng 2 nhưng lại không xuất hiện ở cột Phân xưởng 1 sẽ được bôi màu cam.


![](https://hddtvn.com/wp-content/uploads/2021/01/92-1024x545-1.png)


Từ việc thể hiện bằng màu sắc phần dữ liệu khác nhau của 2 cột dữ liệu trong Excel giúp chúng ta so sánh dữ liệu một cách nhanh chóng, dễ dàng, hiệu quả và tiết kiệm thời gian hơn so với làm bằng các cách thủ công bình thường.


*Như vậy, bài viết trên hddtvn đã hướng dẫn các bạn cách so sánh 2 cột dữ liệu trên Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreKhi muốn so sánh 2 cột dữ liệu thì đa phần các bạn sẽ so…

