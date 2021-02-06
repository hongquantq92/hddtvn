6 Cách khắc phục lỗi Excel không hiện bảng tính cho dân kế toán
===============================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/6-cach-khac-phuc-loi-excel-khong-hien-bang-tinh-cho-dan-ke-toan/)Trong quá trình làm việc, kế toán đôi khi gặp phải lỗi Excel không hiện bảng tính, mà chỉ có một cửa sổ trống màu xám. Chắc hẳn điều này đã không ít lần khiến bạn cảm thấy hoảng hốt, bực mình. Nó xảy ra và làm tiêu tốn thời gian làm việc của bạn …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Trong quá trình làm việc, kế toán đôi khi gặp phải lỗi Excel không hiện bảng tính, mà chỉ có một cửa sổ trống màu xám. Chắc hẳn điều này đã không ít lần khiến bạn cảm thấy hoảng hốt, bực mình. Nó xảy ra và làm tiêu tốn thời gian làm việc của bạn trong khi còn một khối lượng công việc tương đối lớn đang chờ phải giải quyết. Bạn cũng không biết nguyên nhân là gì và cần phải xử lí ra sao. Bài viết dưới đây sẽ bật mí cho bạn 6 cách để *khắc phục lỗi Excel* này.


![cách sửa lỗi excel không hiện bảng tính](https://hddtvn.com/wp-content/uploads/2021/01/cach-sua-loi-excel-khong-hien-bang-tinh.png)


Cách 1: Bật DDE (Dynamic Data Exchange)
---------------------------------------


**Nguyên nhân**


Lỗi Excel bảng tính không hiện có thể là do hộp kiểm Ignore other applications that use Dynamic Data Exchange (DDE) trên trang Options của Excel đang được chọn.


Bạn tùy chọn Ignore -> Excel sẽ bỏ qua các thông điệp DDE được gửi đến nó bởi các chương trình khác -> Thông điệp DDE do Window Explorer gửi đến Excel sẽ bị bỏ qua -> Excel sẽ không mở file bảng tính bạn đã chọn.


**Cách khắc phục**


Trước hết, bạn nhấp đôi chuột lên một file bảng tính Excel trên Windows Explorer -> Một thông điệp DDE được gửi đến Excel. Thông điệp này sẽ báo cho Excel biết rằng nó cần mở file bảng tính bạn vừa nhấp đôi chuột.


Sau đó, bạn tiếp tục làm theo các bước sau để điều chỉnh thiết lập:


– **Bước 1**: Mở Microsoft Excel


– **Bước 2**: Trên trình đơn File, chọn Options


– **Bước 3**: Nhấp vào thẻ Advanced


– **Bước 4**: Kéo xuống đến General -> Bỏ dấu tích chọn tại hộp kiểm Ignore other applications that use Dynamic Data Exchange (DDE) nếu đang được chọn -> Nhấn OK.


![](https://hddtvn.com/wp-content/uploads/2021/01/loi-excel-khong-hien-bang-tinh5.png)


Cách 2: Bỏ ẩn bảng tính trong thẻ View
--------------------------------------


**Nguyên nhân**


Bạn có thể đang để tính năng ẩn bảng tính trên trình đơn View mà không biết, dẫn đến bảng tính của bạn không thể hiện thị khi mwor file Excel.


**Cách khắc phục**


– **Bước 1**: Mở trình đơn View (vì View có thể cho phép ẩn bảng tính, khiến bạn không thể mở được bảng tính)


– **Bước 2**: Kiểm tra tùy chọn xem đã ở trạng thái Unhide chưa.


![lỗi excel không hiên bảng tính](https://hddtvn.com/wp-content/uploads/2021/01/loi-excel-khong-hien-bang-tinh.jpg)


Cách 3: Tắt Add-in
------------------


**Nguyên nhân**


Lỗi Excel bảng tính không hiện có thể là do Excel add-in và COM add-in gây ra. Vì hai loại add-in nằm ở hai thư mục khác nhau, để kiểm tra và tắt, bạn cần tắt lần lượt từng add-in.


**Cách khắc phục**


Bạn có thể làm theo các bước dưới đây:


– **Bước 1**: Trên trình đơn File -> Chọn Options -> Chọn thẻ Add-ins


– **Bước 2**: Dưới cùng màn hình có danh sách Manage -> Chọn COM Add-ins -> Nhấn nút Go


– **Bước 3**: Bỏ chọn một trong những add-in trong danh sách -> Nhân nút OK


– **Bước 4**: Khởi động lại Excel


![lỗi excel không hiện bảng tính](https://hddtvn.com/wp-content/uploads/2021/01/loi-excel-khong-hien-bang-tinh2.png)


*Chú ý:*


*– Nếu vấn đề vẫn chưa được giải quyết, hãy làm lại theo các bước trên, trừ bước 3.*


*– Nếu gặp vấn đề sau khi đã tắt tất cả các COM Add-in, bạn hãy lặp lại trình tự các bước ở trên, trừ bước 2. Sau đó, thử tắt lần lượt các Excel add-in như ở bước 3.*


>> [Những hàm excel phổ biến nhất trong kế toán hiện nay](#)


Cách 4: Thiết lập lại liên kết tập tin
--------------------------------------


– **Bước 1**: Vào Settings


– **Bước 2**: Chọn Apps -> Chọn Default apps.


![lỗi excel không hiện bảng tính](https://hddtvn.com/wp-content/uploads/2021/01/loi-excel-khong-hien-bang-tinh3.jpg)


Cách 5: Vô hiệu hóa chức năng tăng tốc phần cứng
------------------------------------------------


**Nguyên nhân**


Bảng tính trong Excel có thể không hiện là do chức năng tăng tốc phần cứng (hardware acceleration).


**Cách khắc phục**


Để giải quyết vấn đề này, bạn nên tạm thời tắt chức năng này cho đến khi nhà sản xuất card màn hình phát hành bản vá.


Bạn nên thực hiện thao tác theo các bước dưới đây:


– **Bước 1**: Chọn vào Files -> Options -> Chọn thẻ Advanced từ khung bên trái


– **Bước 2**: Di chuyển đến nhóm Display -> Đánh dấu chọn vào hộp kiểm Disable hardware graphics acceleration.


![lỗi excel không hiện bảng tính](https://hddtvn.com/wp-content/uploads/2021/01/loi-excel-khong-hien-bang-tinh4.jpg)


Cách 6: Sửa chữa lại Office
---------------------------


Cách này chỉ nên được áp dụng khi bạn đã thử tất cả những cách trên mà vẫn thất bại. Khi đó, bạn cần phải chạy trình sửa chữa bộ ứng dụng Office (Repair Office), online hay offline đều được. Cách này sẽ giúp sửa các vấn đề xảy ra với bộ ứng dụng Office của bạn.


Nếu sau khi chạy Repair Office, bạn vẫn gặp vấn đề, bạn cần gỡ sạch Microsoft Office (Clean Uninstall). Sau đó, bạn sẽ cài đặt lại bộ ứng dụng văn phòng đến từ Microsoft.


Có thể sẽ còn nhiều cách khác để khắc phục lỗi không hiển thị bảng tính trên Excel, nhưng đây là 6 cách phổ biến, đơn giản và hiệu quả mà bài viết muốn chia sẻ đến bạn. Hy vọng sau khi áp dụng thực hành những cách trên, các bạn nhân viên kế toán sẽ không còn bối rối và không biết cách xử lí khi gặp phải tình trạng này nữa.


>> [Bật mí 10 thủ thuật Excel mà mọi kế toán viên cần biết](#)








































**5**  

/  

**5**  

(  

**1**  

  

 bình chọn   

)


moreTrong quá trình làm việc, kế toán đôi khi gặp phải lỗi Excel không hiện…

