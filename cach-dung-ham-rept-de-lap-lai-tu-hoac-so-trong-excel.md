Cách dùng hàm REPT để lặp lại từ hoặc số trong Excel
====================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-dung-ham-rept-de-lap-lai-tu-hoac-so-trong-excel/)Với hàm REPT, các ký tự trong bảng tự động lặp lại theo số lần mà bạn muốn, thay vì phải nhập thủ công lần lượt từng ký tự. Hàm REPT sẽ trả về chuỗi văn bản, lặp lại ký tự hoặc số theo đúng số lần trong công thức mà bạn nhập. Như vậy …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Với hàm REPT, các ký tự trong bảng tự động lặp lại theo số lần mà bạn muốn, thay vì phải nhập thủ công lần lượt từng ký tự. Hàm REPT sẽ trả về chuỗi văn bản, lặp lại ký tự hoặc số theo đúng số lần trong công thức mà bạn nhập. Như vậy với hàm REPT, chúng ta tiết kiệm thời gian làm việc hơn rất nhiều. Bài viết này sẽ chia sẻ với các bạn cách sử dụng hàm REPT để lặp lại từ hoặc số trong Excel nhé.**


### 1. Cấu trúc hàm REPT


Cú pháp hàm: **=REPT(text; number\_times)**


*Trong đó:*




* **Text**: đối số bắt buộc, là văn bản mà bạn muốn lặp lại.

* **Number\_times**: đối số bắt buộc, là một số dương chỉ rõ số lần lặp lại văn bản.



*Lưu ý:*




* Nếu **number\_times** bằng 0 (không), thì hàm REPT trả về “” (văn bản trống).

* Nếu **number\_times** không phải là số nguyên thì nó bị cắt cụt.

* Nếu **number\_times** là số âm hoặc không phải là số thì hàm trả về lỗi #VALUE!.

* Kết quả của hàm REPT không được dài quá 32.767 ký tự, nếu không, hàm REPT trả về lỗi #VALUE!.



### 2. Cách sử dụng hàm REPT


#### a. Sử dụng hàm REPT để lặp lại từ hoặc số


Ví dụ, để lặp lại tên mặt hàng đầu tiên 2 lần thì ta có công thức sau: **=REPT(A2;2)**


Hoặc: **=REPT(“Harlech”;2)**


![](https://hddtvn.com/wp-content/uploads/2021/01/au0NzcU.png)


Trong trường hợp đối số **number\_times** là số âm thì hàm sẽ trả về giá trị lỗi #VALUE!.


![](https://hddtvn.com/wp-content/uploads/2021/01/IwNSJRz.png)


#### b. Sử dụng hàm REPT để vẽ biểu đồ


Hàm REPT còn có một ứng dụng khá thú vị đó chính là vẽ biểu đồ. Đầu tiên, các bạn nhập công thức sau vào ô F2:


**=REPT(“g”;E2/$E$11*100)**


Công thức có nghĩa là lặp lại số lần từ **g** số lần tương ứng với tỷ lệ số tiền của mặt hàng đầu tiên trên tổng số tiền của tất cả mặt hàng.


Tiếp theo các bạn sao chép công thức cho tất cả ô còn lại rồi chọn Font chữ **Webdings**. Khi sử dụng font chữ này thì từ **g** sẽ tương ứng với các ký hiệu đặc biệt.


Chỉ cần như vậy là biểu đồ tỷ trọng của các mặt hàng đã được tạo một cách nhanh chóng.


![](https://hddtvn.com/wp-content/uploads/2021/01/lSZDCYv.png)


Tiếp theo, các bạn có thể tô màu cho biểu đồ của mình trông nổi bật hơn tại mục Font Color.


[![](https://hddtvn.com/wp-content/uploads/2021/01/Q9oKBSm.png)](https://hddtvn.com/wp-content/uploads/2021/01/Q9oKBSm.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm REPT để lặp lại từ hoặc số trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreVới hàm REPT, các ký tự trong bảng tự động lặp lại theo số lần…

