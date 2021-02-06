Cách sử dụng hàm PRODUCT để thực hiện phép tính nhân trong Excel
================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-product-de-thuc-hien-phep-tinh-nhan-trong-excel/)Hàm PRODUCT là một trong những hàm toán học được sử dụng phổ biến trong Excel. Đây là hàm dùng để nhân trong Excel, nó sẽ nhân tất cả các đối số được đưa vào và trả về tích của chúng. Trong bài viết sau đây, hddtvn sẽ giới thiệu đến bạn đọc cách sử …
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm PRODUCT là một trong những hàm toán học được sử dụng phổ biến trong Excel. Đây là hàm dùng để nhân trong Excel, nó sẽ nhân tất cả các đối số được đưa vào và trả về tích của chúng.** **Trong bài viết sau đây, hddtvn sẽ giới thiệu đến bạn đọc cách sử dụng hàm PRODUCT để thực hiện phép tính tích các giá trị.**


### 1. Cấu trúc hàm PRODUCT


Cú pháp hàm: **=PRODUCT(number1; [number2]; …)**


*Trong đó:*




* **number1**: đối số bắt buộc, là số hoặc phạm vi thứ nhất mà bạn muốn nhân.

* **number2**: đối số tùy chọn, là các số hoặc phạm vi bổ sung mà bạn muốn nhân tiếp theo. Tối đa 255 đối số.



*Lưu ý:*




* Hàm PRODUCT nhân tất cả các đối số đã cho với nhau và trả về tích của chúng.

* Bạn cũng có thể thực hiện thao tác này bằng cách dùng toán tử toán học của phép nhân (*).

* Hàm PRODUCT hữu ích khi bạn cần nhân nhiều ô với nhau.

* Nếu đối số là mảng hoặc tham chiếu, thì chỉ có các số trong mảng hoặc tham chiếu đó mới được nhân. Các ô trống, giá trị logic và văn bản trong mảng hoặc tham chiếu bị bỏ qua.



### 2. Cách sử dụng hàm PRODUCT


Ví dụ ta có bảng dữ liệu như hình và cần nhân theo các yêu cầu như sau:




* Nhân các số trong các ô từ A2 đến D2

* Nhân các số trong các ô từ A2 đến D2 bằng toán tử (*)

* Nhân các số trong các ô từ A2 đến A5 rồi nhân kết quả đó với 3

* Nhân các số trong các ô từ A2 đến D2 và từ A4 đến D4

* Nhân các số trong các ô từ A2 đến D5



![](https://hddtvn.com/wp-content/uploads/2021/01/hmEgLnS.png)


Đầu tiên, để nhân các số trong các ô từ A2 đến D2. Áp dụng cấu trúc hàm như trên, ta có công thức tính như sau:


**=PRODUCT(A2:D2)**


Và với yêu cầu nhân các số trong các ô từ A2 đến D2 bằng toán tử (*), ta có công thức tính như sau:


**=A2*B2*C2*D2**


Ta có thể thấy 2 công thức này cho ra kết quả giống nhau như hình dưới. Do hàm Product thực hiện theo toán tử (*) nhưng sẽ tiện dụng hơn rất nhiều.


![](https://hddtvn.com/wp-content/uploads/2021/01/dWMG7Kk.png)


Áp dụng cấu trúc hàm, ta sẽ có công thức tính cho các yêu cầu còn lại như sau:




* Nhân các số trong các ô từ A2 đến A5 rồi nhân kết quả đó với 3: **=PRODUCT(A2:A5;3)**

* Nhân các số trong các ô từ A2 đến D2 và từ A4 đến D4: **=PRODUCT(A2:D2;A4:D4)**

* Nhân các số trong các ô từ A2 đến D5: **=PRODUCT(A2:D5)**



![](https://hddtvn.com/wp-content/uploads/2021/01/8ayjgu4.png)


*Hy vọng qua bài viết này, các bạn đã có thể hiểu rõ được cách sử dụng hàm PRODUCT để thực hiện phép tính nhân trong Excel. Chúc các bạn thành công!*


moreHàm PRODUCT là một trong những hàm toán học được sử dụng phổ biến trong…

