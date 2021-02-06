Cách sử dụng hàm MOD để lấy số dư của phép chia trong Excel
===========================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-mod-de-lay-so-du-cua-phep-chia-trong-excel/)Trong các bảng tính số liệu trên bảng tính Excel, một số bài toán cần dùng đến số dư của phép chia để kết hợp với nhiều hàm khác để thực hiện tính toán. Để lấy được phần dư của phép chia trong Excel chúng ta sử dụng Hàm MOD. Bài viết dưới đây sẽ …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Trong các bảng tính số liệu trên bảng tính Excel, một số bài toán cần dùng đến số dư của phép chia để kết hợp với nhiều hàm khác để thực hiện tính toán. Để lấy được phần dư của phép chia trong Excel chúng ta sử dụng Hàm MOD. Bài viết dưới đây sẽ giúp bạn hiểu rõ hơn về cú pháp cũng như cách sử dụng hàm MOD trong Excel thông qua các ví dụ cụ thể.**


### 1. Cấu trúc hàm MOD


Cú pháp hàm: **=MOD (number; divisor)**


*Trong đó:*




* **Number**: đối số bắt buộc, là số bị chia.

* **Divisor**: đối số bắt buộc, là số chia.



*Lưu ý:*




* Kết quả trả về sẽ cùng dấu với **divisor**

* Nếu **divisor** số là 0 thì hàm MOD trả về giá trị lỗi #DIV/0! .



### 2. Cách sử dụng hàm MOD


Ví dụ nếu ta cần lấy số dư của phép tính 5 chia 3. Áp dụng cấu trúc hàm MOD như trên ta có công thức tính như sau: **=MOD(5;3)**


Ta sẽ thu được kết quả là 2. Do 5 = 2 + (3 x 1)


![](https://hddtvn.com/wp-content/uploads/2021/01/5Mk1wKC.png)


Nếu ta cần lấy số dư của -5 chia cho 3. Ta có công thức như sau: **=MOD(-5;3)**


Kết quả của phép tính trên sẽ là 1 do kết quả sẽ là số dương cùng dấu với số chia là 3. Trong trường hợp này phép tính để lấy được kết quả số dư sẽ là: -5 = (-2) x 3 + 1.


![](https://hddtvn.com/wp-content/uploads/2021/01/C5w2k0y.png)


Nếu ta cần lấy số dư của 5 chia cho -3. Ta có công thức như sau: **=MOD(5;-3)**


Kết quả của phép tính trên sẽ là -1 do kết quả sẽ là số âm cùng dấu với số chia là -3. Trong trường hợp này phép tính để lấy được kết quả số dư sẽ là: 5 = (-2) x -3 – 1.


![](https://hddtvn.com/wp-content/uploads/2021/01/R1piCDR.png)


Nếu ta cần lấy số dư của -5 chia cho -3. Ta có công thức như sau: **=MOD(-5;-3)**


Kết quả của phép tính trên sẽ là -2 do kết quả sẽ là số âm cùng dấu với số chia là -3. Trong trường hợp này phép tính để lấy được kết quả số dư sẽ là: -5 = -3 – 2.


![](https://hddtvn.com/wp-content/uploads/2021/01/mAmonDj.png)


Nếu ta cần lấy số dư của 5 chia cho 0. Ta có công thức như sau: **=MOD(5;0)**


Lúc này hàm sẽ trả về lỗi #DIV/0! vì trong toán học không thể chia một số cho 0 được.


![](https://hddtvn.com/wp-content/uploads/2021/01/Tla5090.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm MOD để lấy số dư của phép chia trong Excel. Chúc các bạn thành công!*


moreTrong các bảng tính số liệu trên bảng tính Excel, một số bài toán cần…

