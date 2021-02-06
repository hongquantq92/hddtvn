Hướng dẫn cách nối các ô với nhau trong Excel
=============================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-cach-noi-cac-o-voi-nhau-trong-excel/)Trong quá trình làm việc trên Excel, nhiều lúc bạn sẽ cần nối ký tự trong các ô khác nhau vào trong 1 ô. Bài viết sau đây sẽ hướng dẫn các bạn các cách nối các ô với nhau trong Excel. Sử dụng hàm CONCATENATE để nối các ô với nhau Cú pháp hàm: …
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Trong quá trình làm việc trên Excel, nhiều lúc bạn sẽ cần nối ký tự trong các ô khác nhau vào trong 1 ô. Bài viết sau đây sẽ hướng dẫn các bạn các cách nối các ô với nhau trong Excel.**


### Sử dụng hàm CONCATENATE để nối các ô với nhau


Cú pháp hàm: **=CONCATENATE(text1;[text2];…)**


*Trong đó:*




* text1 là đối số bắt buộc, là chuỗi đầu tiên cần ghép nối nó có thể là giá trị văn bản, số hoặc tham chiếu ô.

* text2 là đối số tùy chọn, là chuỗi thứ hai cần ghép nối có thể có tối đa 255 mục (8192 ký tự).



*Lưu ý:*




* Hàm CONCATENATE yêu cầu ít nhất một “text” đối số để làm việc.

* Các ký tự đặc biệt, ngày tháng… cần để trong dấu ngoặc kép ” “.

* Trong một công thức CONCATENATE, bạn có thể nối ghép lên tới 255 chuỗi, tổng cộng là 8,192 ký tự.

* Kết quả của hàm CONCATENATE luôn luôn là một chuỗi văn bản, ngay cả khi tất cả các giá trị nguồn là các con số.

* CONCATENATE không nhận biết các mảng. Mỗi tham chiếu ô phải được liệt kê riêng. Ví dụ, bạn nên viết =CONCATENATE(A1, A2, A3) thay vì =CONCATENATE(A1:A3).

* Nếu ít nhất một trong các đối số của hàm CONCATENATE không hợp lệ, công thức sẽ trả về lỗi #VALUE!.



Ví dụ: Sử dụng hàm CONCATENATE để nối hai ô


[![](https://hddtvn.com/wp-content/uploads/2021/01/Q1myDO0.png)](https://hddtvn.com/wp-content/uploads/2021/01/Q1myDO0.png)


Sử dụng hàm CONCATENATE để nối chuỗi và thêm khoảng cách giữa các chuỗi nối


![](https://hddtvn.com/wp-content/uploads/2021/01/GxtE3hc.png)


Sử dụng hàm CONCATENATE để nối chuỗi và thêm chữ vào chuỗi


![](https://hddtvn.com/wp-content/uploads/2021/01/iGmq2M7.png)


### Sử dụng toán tử “&” để nối các ô trong Excel


Toán tử & là một cách khác để nối các ô trong Excel. Sự khác biệt duy nhất giữa hàm CONCATENATE và “&” là hàm CONCATENATE có giới hạn 255 còn & thì không có. Ngoài ra, không có gì khác nhau giữa hai phương pháp nối này, cũng không có sự khác biệt về hiệu lực giữa các công thức CONCATENATE và “&”.


Cách sử dụng toán tử & để nối chuỗi, ta thực hiện viết hàm như sau: **=B3&C3**


![](https://hddtvn.com/wp-content/uploads/2021/01/83ux9BV.png)


Ta có thể nối thêm khoảng cách vào các chuỗi. Ví dụ: **=B3&” “&C3**


![](https://hddtvn.com/wp-content/uploads/2021/01/kDRUIBj.png)


Hoặc có thể thêm các ký tự vào chuỗi như sau: **=B3&C3&” chia sẻ kiến thức”**


![](https://hddtvn.com/wp-content/uploads/2021/01/zxTy0th.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách nối các ô với nhau trong Excel. Chúc các bạn thành công!*


moreTrong quá trình làm việc trên Excel, nhiều lúc bạn sẽ cần nối ký tự…

