Hướng dẫn lập bộ BCTC theo Thông tư 133/2016/TT-BTC trên HTKK
=============================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-lap-bo-bctc-theo-thong-tu-133-2016-tt-btc-tren-htkk/)Bạn không biết cách lập báo cáo tài chính theo Thông tư 133 trên phần mềm HTKK như thế nào để đảm bảo đúng quy định pháp luật? Vậy đừng bỏ sót từng từ trong bài viết này nhé. Bởi hddtvn sẽ hướng dẫn các bạn cách lập bộ BCTC theo Thông tư 133/2016/TT-BTC trên …
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Bạn không biết cách lập báo cáo tài chính theo Thông tư 133 trên phần mềm HTKK như thế nào để đảm bảo đúng quy định pháp luật? Vậy đừng bỏ sót từng từ trong bài viết này nhé. Bởi hddtvn sẽ hướng dẫn các bạn cách lập bộ BCTC theo Thông tư 133/2016/TT-BTC trên phần mềm HTKK một cách chi tiết nhất. Mời các bạn theo dõi.**


### Bước 1: Đăng nhập vào phần mềm


Đầu tiên, các bạn cần mở phần mềm HTKK lên, sau đó nhập mã số thuế của đơn vị cần lập bộ báo cáo tài chính theo thông tư 133 rồi bấm **Đồng ý**.


![](https://hddtvn.com/wp-content/uploads/2021/01/tT2sql8.png)


### Bước 2: Chọn báo cáo tài chính


Sau khi màn hình chính của phần mềm hiện lên, cho mục **Kê khai** => **Báo cáo tài chính** => **Bộ báo cáo tài chính (B01a – DNN)(Thông tư 133/2016/TT-BTC)**. Lưu ý đây là báo cáo tài chính dành cho doanh nghiệp đáp ứng giả định hoạt động liên tục.


![](https://hddtvn.com/wp-content/uploads/2021/01/50a5l73.png)


### Bước 3: Chọn kỳ tính thuế


Sau khi chọn bộ báo cáo tài chính, cửa sổ **Chọn kỳ tính thuế** hiện lên. Tại đây các bạn có thể chọn năm tính thuế và các phụ lục kê khai.


Lưu ý chỉ được chọn 1 trong 2 bảng lưu chuyển tiền tệ là LCTTTT (trực tiếp) và LCTTGT (gián tiếp).


![](https://hddtvn.com/wp-content/uploads/2021/01/IO3rpjc.png)


### Bước 4: Lập bộ BCTC


#### a. Lập Báo cáo tình hình tài chính (Mẫu số B01a – DNN)




* Ngày tháng năm: hệ thống tự động hiển thị

* Người nộp thuế, Mã số thuế, Tên đại lý thuế (nếu có): Tự động hiển thị theo thông tin chung đã nhập tại menu Hệ thống/ Thông tin doanh nghiệp/ Người nộp thuế.

* Hỗ trợ lấy dữ liệu từ năm trước: mặc định không check. Nếu tích chọn ứng dụng hỗ trợ lấy dữ liệu từ cột <Số cuối năm> của năm trước liền kề sang cột <Số đầu năm> của năm nay.

* Các cột Chỉ tiêu, Mã số: Hiển thị theo tên các chỉ tiêu và mã số của Báo cáo tình hình tài chính năm (thuộc bộ BCTC mẫu số B01a – DNN) dành cho doanh nghiệp nhỏ và vừa đáp ứng giả định hoạt động liên tục theo Thông tư số 133/2016/TT-BTC.

* Cột Thuyết minh: Cho phép sửa, tối đa 50 ký tự. Ô checkbox “Tích chọn để nhập cột Thuyết minh”, mặc định là Không tích, nếu không thì khóa không cho nhập, tích chọn thì mở khóa cho nhập cột Thuyết Minh.

* Số cuối năm, Số đầu năm: cho phép nhập kiểu số tối đa 15 ký tự.

* Tất cả các chỉ tiêu ban đầu nhận giá trị mặc định là 0.

* Các chỉ tiêu :110, 121, 122, 123, 131, 132, 133, 134, 135, 141, 151,161, 170, 181, 182, 311, 312, 313, 314, 315, 316, 317, 411, 413, 416 => Nhập kiểu số >=0

* Các chỉ tiêu: 124, 136, 142, 152, 162, 414 Nhập <= 0 hiển thị (…)

* Các chỉ tiêu: 318, 319, 320, 412, 415, 417: Nhập âm dương

* Các chỉ tiêu có công thức sẽ tự nhảy và không tự sửa.



![](https://hddtvn.com/wp-content/uploads/2021/01/vNf2u4I.png)


#### b. Lập Báo cáo kết quả hoạt động kinh doanh




* 01, 21, 31, 32, 51: Nhập kiểu số >=0

* 02: Nhập kiểu số >=0. Kiểm tra chỉ tiêu [02] <= chỉ tiêu [01]. Nếu không thỏa mãn cảnh báo vàng

* 10: Hỗ trợ tính theo công thức chỉ tiêu [10]= chỉ tiêu [01] – chỉ tiêu [02], không cho sửa, tổng có thể âm

* 11, 22: Nhập kiểu số âm dương

* 20: Hỗ trợ tính theo công thức chỉ tiêu [20]= chỉ tiêu [10]-chỉ tiêu [11], không cho sửa, tổng có thể âm

* 23: Nhập kiểu số âm dương, kiểm tra chỉ tiêu [23] <= chỉ tiêu [22], cảnh báo đỏ nếu không thỏa mãn

* 24: Nhập âm dương

* 30: Hỗ trợ tính theo công thức [30]= [20] + [21] – [22] – [24], không cho sửa, tổng có thể âm

* 40: Hỗ trợ tính theo công thức [40] = [31] – [32], không cho sửa, tổng có thể âm

* 50: Hỗ trợ tính theo công thức [50] = [30] + [40], không cho sửa, tổng có thể âm

* 60: Hỗ trợ tính theo công thức [60]=[50] – [51], không cho sửa, tổng có thể âm



[![](https://hddtvn.com/wp-content/uploads/2021/01/3D25E4p.png)](https://hddtvn.com/wp-content/uploads/2021/01/3D25E4p.png)


#### c. Lập Lưu chuyển tiền tệ trực tiếp




* Tất cả các chỉ tiêu ban đầu nhận giá trị mặc định là 0.

* Các cột phải nhập: cột [Năm nay], [Năm trước]

* Ứng dụng hỗ trợ lấy dữ liệu cột [Năm trước] của năm trước sang cột [Năm trước] của năm nay, cho sửa. Nếu sửa khác thì ứng dụng cảnh báo vàng “Số đầu năm của năm nay khác với Số cuối năm của năm trước”

* Cột “Thuyết minh”: cho phép nhập tối đa 50 ký tự.

* Các chỉ tiêu 01, 06, 24, 25, 31, 33, 60: Nhập kiểu số >=0, mặc định là 0

* Các chỉ tiêu 02, 03, 04, 05, 07, 21, 23, 32, 34, 35: Nhập <= 0 hiển thị (…), mặc định là 0

* Các chỉ tiêu: 22, 61 Nhập âm dương

* Các chỉ tiêu có công thức sẽ tự nhảy và không tự sửa.



![](https://hddtvn.com/wp-content/uploads/2021/01/qyfAHjl.png)


#### d. Lập Lưu chuyển tiền tệ gián tiếp




* 01: Nhập âm, dương

* 02: Hỗ trợ tính [02]=[03]+[04]+[05]+[06]+[07]+[08], không cho sửa, tổng có thể âm

* 03, 07: Nhập >=0

* 04, 05, 06, 08: Nhập âm dương

* 09: Hỗ trợ tính theo công thức [09]=[10]+[11]+[12]+[13]+[14]+[15]+[16]+[17]+[18], không cho sửa, tổng có thể âm

* 10, 11, 12, 13, 14:Nhập âm, dương

* 15, 16, 18: Nhập kiểu số <=0

* 17: Nhập kiểu số >=0

* 20: Hỗ trợ tính theo công thức [20]=[01]+[02]+[09], không cho sửa, tổng có thể âm

* 21, 23: Nhập <=0, hiển thị (…)

* 22: Nhập âm dương

* 24, 25: Nhập kiểu số >=0

* 30: Hỗ trợ tính theo công thức [30]=[21]+[22]+[23]+[24]+[25], không cho sửa, tổng có thể âm

* 31, 33: Nhập kiểu số >=0

* 32, 34, 35: Nhập <=0, hiển thị (…)

* 40: Hỗ trợ tính theo công thức [40]=[31]+[32]+[33]+[34]+[35], không cho sửa, tổng có thể âm

* 50: Hỗ trợ tính theo công thức [50]=[20]+[30]+[40], không cho sửa, tổng có thể âm

* 60: Nhập kiểu số >=0

* 61: Nhập âm dương

* 70: Hỗ trợ tính theo công thức [70] = [50]+[60]+[61], không cho sửa, tổng có thể âm



![](https://hddtvn.com/wp-content/uploads/2021/01/d5oI1aS.png)


Sau khi đã nhập xong chỉ tiêu trên tất cả các báo cáo, tiến hành nhập tên Người lập biểu, Kế toán trưởng, Giám đốc và Ngày ký. Sau đó bấm **Ghi** để lưu bộ báo cáo và bấm **Kết xuất** để xuất bộ báo cáo ra Excel hoặc file XML nếu cần.


*Chúc các bạn thành công!*


#### 


moreBạn không biết cách lập báo cáo tài chính theo Thông tư 133 trên phần…

