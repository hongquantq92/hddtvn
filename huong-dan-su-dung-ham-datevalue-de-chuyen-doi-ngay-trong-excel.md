Hướng dẫn sử dụng hàm DATEVALUE để chuyển đổi ngày trong Excel
==============================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-su-dung-ham-datevalue-de-chuyen-doi-ngay-trong-excel/)Hàm DATEVALUE có được sử dụng để chuyển đổi ngày được lưu trữ ở dạng văn bản sang dạng số seri mà Excel công nhận là ngày tháng. Hãy theo dõi bài viết sau để hiểu rõ hơn về cách sử dụng hàm DATEVALUE trong Excel. 1. Cấu trúc hàm DATEVALUE Cú pháp hàm: =DATEVALUE(date\_text) …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm DATEVALUE có được sử dụng để chuyển đổi ngày được lưu trữ ở dạng văn bản sang dạng số seri mà Excel công nhận là ngày tháng. Hãy theo dõi bài viết sau để hiểu rõ hơn về cách sử dụng hàm DATEVALUE trong Excel.**


### 1. Cấu trúc hàm DATEVALUE


Cú pháp hàm: **=DATEVALUE(date\_text)**


*Trong đó:*




* **Date\_text**: đối số bắt buộc, là ngày tháng được viết dưới dạng văn bản hoặc tham chiếu đến ô chứa ngày tháng dạng văn bản.



*Lưu ý:*




* **Date\_text** là đối số bắt buộc có, chỉ có thể bỏ qua phần năm của đối số. Tức là chỉ cần nhập ngày tháng còn năm thì hàm DATEVALUE sẽ mặc định là năm hiện tại.

* **Date\_text** phải để trong ngoặc kết biểu hiện giá trị ngày. Nếu không hàm sẽ cho giá trị lỗi là #VALUE.



### 2. Cách sử dụng hàm DATEVALUE


Ví dụ để chuyển ngày 11/4/2019 sang định dạng seri ta có công thức như sau:


**=DATEVALUE(“11/4/2019”)**


![](https://hddtvn.com/wp-content/uploads/2021/01/JXU6vOy.png)


Nếu trong trường hợp sử dụng hàm DATEVALUE tham chiếu đến ô có format là Date mà không phải dạng văn bản thì hàm sẽ trả về lỗi #VALUE!.


![](https://hddtvn.com/wp-content/uploads/2021/01/dUfLcQz.png)


Trong trường hợp này, ta cần sửa định dạng ngày tháng của ô A4 thành văn bản bằng cách thêm:


**=”11/04/2019″**


Lúc này ô A4 sẽ được coi là dạng văn bản và hàm DATEVALUE sẽ hoạt động.


![](https://hddtvn.com/wp-content/uploads/2021/01/GUuWr4G.png)


Trong trường hợp giá trị năm bị bỏ qua thì hàm sẽ coi là năm hiện tại.


![](https://hddtvn.com/wp-content/uploads/2021/01/DoPORya.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm DATEVALUE trong Excel. Hy vọng qua bài viết này, các bạn đã nắm rõ được cách sử dụng hàm này. Chúc các bạn thành công!*


moreHàm DATEVALUE có được sử dụng để chuyển đổi ngày được lưu trữ ở dạng…

