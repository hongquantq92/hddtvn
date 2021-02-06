Cách sử dụng hàm ODD để làm tròn đến số nguyên lẻ gần nhất trong Excel
======================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-odd-de-lam-tron-den-so-nguyen-le-gan-nhat-trong-excel/)Hàm ODD được dùng để làm tròn đến số nguyên lẻ gần nhất. Hàm này thường ứng dụng để làm tròn số lượng, số sản phẩm, số tiền, doanh thu, giá bán… Trong bài viết này, hddtvn sẽ giới thiệu đến các bạn cách sử dụng chi tiết hàm ODD, có kèm ví dụ cụ …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm ODD được dùng để làm tròn đến số nguyên lẻ gần nhất. Hàm này thường ứng dụng để làm tròn số lượng, số sản phẩm, số tiền, doanh thu, giá bán…** **Trong bài viết này, hddtvn sẽ giới thiệu đến các bạn cách sử dụng chi tiết hàm ODD, có kèm ví dụ cụ thể nhé.**


### 1. Cấu trúc hàm ODD


Cú pháp hàm: **=ODD(number)**


*Trong đó:***Number** đối số bắt buộc, là giá trị cần làm tròn.


*Lưu ý:*




* Nếu **number** không phải dạng số, hàm ODD trả về giá trị lỗi #VALUE! .

* Không xét đến dấu của đối số **number**, giá trị được làm tròn ra xa số không.

* Nếu **number** là số nguyên lẻ, thì không làm tròn.



### 2. Cách sử dụng hàm ODD


Ví dụ ta cần làm tròn những số sau đây:




* Làm tròn 4 đến số nguyên lẻ gần nhất

* Làm tròn 4,5 đến số nguyên lẻ gần nhất

* Làm tròn 3,4 đến số nguyên lẻ gần nhất

* Làm tròn -4 đến số nguyên lẻ gần nhất

* Làm tròn -4,5 đến số nguyên lẻ gần nhất

* Làm tròn -3,4 đến số nguyên lẻ gần nhất



![](https://hddtvn.com/wp-content/uploads/2021/01/HgwYhXr.png)


Áp dụng cấu trúc hàm như trên, ta có công thức tính cho các trường hợp như sau:




* Làm tròn 4 đến số nguyên lẻ gần nhất: **=ODD(4)**

* Làm tròn 4,5 đến số nguyên lẻ gần nhất: **=ODD(4,5)**

* Làm tròn 3,4 đến số nguyên lẻ gần nhất: **=ODD(3,4)**

* Làm tròn -4 đến số nguyên lẻ gần nhất: **=ODD(-4)**

* Làm tròn -4,5 đến số nguyên lẻ gần nhất: **=ODD(-4,5)**

* Làm tròn -3,4 đến số nguyên lẻ gần nhất: **=ODD(-3,4)**



![](https://hddtvn.com/wp-content/uploads/2021/01/gQfpGKY.png)


Nhìn vào kết quả, ta có thể thấy hàm ODD sẽ làm tròn số ra xa số 0.




* Hàm sẽ làm tròn số 4 thành 5 chứ không phải thành 3 dù 3 và 5 đều gần 4 như nhau. Tương tự với -4 sẽ được làm tròn thành -5 chứ không phải -3.

* Với 4,5 thì gần với 5 nhất nên sẽ được làm tròn thành 5. Tương tự với -4,5 được làm tròn thành -5.

* Còn với 3,4 thì sẽ được hàm làm tròn thành 5 chứ không phải 3 dù 3,4 gần với 3 hơn. Do hàm sẽ làm tròn ra xa số 0 chứ không phải gần số 0. Tương tự với -3,4 sẽ được làm tròn thành -5.



*Hy vọng qua bài viết trên, các bạn đã hiểu được cách sử dụng hàm ODD để làm tròn trong Excel. Chúc các bạn thành công!*


moreHàm ODD được dùng để làm tròn đến số nguyên lẻ gần nhất. Hàm này…

