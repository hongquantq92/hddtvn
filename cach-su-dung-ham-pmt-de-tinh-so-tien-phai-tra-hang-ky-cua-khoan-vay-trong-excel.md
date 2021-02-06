Cách sử dụng hàm PMT để tính số tiền phải trả hàng kỳ của khoản vay trong Excel
===============================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-pmt-de-tinh-so-tien-phai-tra-hang-ky-cua-khoan-vay-trong-excel/)Hàm PMT là hàm tài chính thường được sử dụng đề tính số phải trả hàng kỳ của một khoản vay có lãi suất không đổi. Bài viết sau đây sẽ hướng dẫn các bạn sử dụng hàm PTM để tính số tiền thanh toán hàng kỳ cho khoản vay trong Excel. 1. Cấu trúc …
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm PMT là hàm tài chính thường được sử dụng đề tính số phải trả hàng kỳ của một khoản vay có lãi suất không đổi. Bài viết sau đây sẽ hướng dẫn các bạn sử dụng hàm PTM để tính số tiền thanh toán hàng kỳ cho khoản vay trong Excel.**


### 1. Cấu trúc hàm PMT


Cú pháp hàm: **=PMT(rate; nper; pv; [fv]; [type])**


*Trong đó:*




* **Rate**: Đối số bắt buộc. Đây là lãi suất hàng kỳ của khoản vay.

* **Nper**: Đối số bắt buộc. Đây là tổng số kỳ thanh toán của khoản vay.

* **Pv**: Đối số bắt buộc. Giá trị hiện tại của khoản vay (nợ gốc).

* **Fv**: Đối số tùy chọn. Giá trị tương lai (hoặc số dư) của khoản tiền sau khi thực hiện việc thanh toán đợt cuối cùng. Nếu không chọn thì hiểu mặc định là sẽ trả hết khoản vay

* **Type**: Đối số tùy chọn. Chọn giữa số 0 hoặc số 1. Nếu Type được bỏ qua thì nó sẽ mặc định là giá trị 0 thể hiện thời điểm thanh toán cuối kì, và giá trị 1 thể hiện thời điểm thanh toán đầu kì.



**Lưu ý:**




* Số tiền thanh toán mà hàm **PMT** trả về bao gồm nợ gốc và lãi nhưng không bao gồm các chi phí, lệ phí, thuế đôi khi đi kèm với khoản vay.

* **Rate** và **Nper** phải thống nhất về cùng 1 giá trị thời gian

* Nếu **Nper** bắt buộc tính theo tháng mà đề bài cho lãi suất theo năm thì phải quy đổi **Rate** về theo tháng



### 2. Cách sử dụng hàm PMT


#### a. Sử dụng hàm PMT để tính toán số tiền phải trả hàng kỳ của khoản vay


Giả sử ta vay 1 khoản trị giá 20 tỷ, kỳ hạn 5 năm với lãi suất là 12%/năm. Tính số tiền phải trả hàng năm của khoản vay.


![](https://hddtvn.com/wp-content/uploads/2021/01/B7FPY85.png)


Để tính số tiền phải trả hàng năm, tại ô D3 ta nhập công thức: **=PMT(C3;B3;A3;0;0)**


Ta thu được kết quả mỗi năm phải trả **3.539.683.283,2** như hình dưới.


![](https://hddtvn.com/wp-content/uploads/2021/01/JR2CCpC.png)


Kết quả của hàm PMT mặc định sẽ là định dạng tiền tệ. Nếu bạn không muốn sử dụng định dạng này ta có thể chỉnh lại định dạng số bằng cách click chuột phải sau đó chọn **Format Cells…** hoặc bạn có thể bấm tổ hợp phím Ctrl + 1 để mở hộp thoại Format Cells


![](https://hddtvn.com/wp-content/uploads/2021/01/4A1QAJp.png)


Hộp thoại **Format Cells** hiện tại, chọn thẻ **Number** => tại mục **Category** ta chọn **Number** => sau đó ta có thể chỉnh các loại định dạng số mong muốn tại đây.


![](https://hddtvn.com/wp-content/uploads/2021/01/HCgkNp7.png)


#### b. Sử dụng hàm PMT để tính số tiền phải gửi tiết kiệm cho đến khi đạt mục tiêu


Giả sử ta đặt ra mục tiêu gửi tiết kiệm định kì hàng tháng trong vòng 6 tháng. Số tiền mong muốn thu về sau 6 tháng là 1 tỷ VND, lãi suất tiền gửi là 1%/tháng.


![](https://hddtvn.com/wp-content/uploads/2021/01/gXEkIe7.png)


Để tính số tiền phải gửi tiết kiệm hàng tháng, ta có công thức tại ô D3: =PMT(C3;B3;0;A3)


Ta thu được kết quả mỗi tháng ta phải gửi 162.548.367 VND để sau 6 tháng ta thu được 1 tỷ VND


![Cách sử dụng hàm PMT](https://hddtvn.com/wp-content/uploads/2021/01/wEfzgWc.png "Cách sử dụng hàm PMT")


*Như vậy, bài viết trên đã hướng dẫn các bạn cách dùng hàm PMT để tính số tiền phải trả hàng kỳ của khoản vay trong Excel. Chúc các bạn thành công!*


moreHàm PMT là hàm tài chính thường được sử dụng đề tính số phải trả…

