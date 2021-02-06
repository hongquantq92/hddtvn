Hướng dẫn cách xóa ảnh hàng loạt trong Microsoft Word
=====================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-cach-xoa-anh-hang-loat-trong-microsoft-word/)Thông thường, để xóa 1 hình ảnh trong Word thì ta chỉ cần nhấn Delete là xong. Vậy với những tài liệu có nhiều hình ảnh mà muốn xóa thì làm thế nào? Nếu xóa thủ công từng ảnh sẽ không khả thi vì tốn nhiều thời gian và thao tác thực hiện. Nếu vẫn chưa …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Thông thường, để xóa 1 hình ảnh trong Word thì ta chỉ cần nhấn Delete là xong. Vậy với những tài liệu có nhiều hình ảnh mà muốn xóa thì làm thế nào? Nếu xóa thủ công từng ảnh sẽ không khả thi vì tốn nhiều thời gian và thao tác thực hiện. Nếu vẫn chưa biết làm thế nào để xóa hàng loạt hình ảnh trong Word thì mời bạn theo dõi bài viết dưới đây của hddtvn nhé**


![Hướng dẫn cách xóa ảnh hàng loạt trong Microsoft Word](https://hddtvn.com/wp-content/uploads/2021/01/xoa-anh-1-1.jpg)


### 1. Xóa ảnh hàng loạt trong Word bằng công cụ Replace


Đầu tiên, các bạn cần mở công cụ Replace bằng cách chọn thẻ **Home** trên thanh công cụ => **Replace**. Hoặc các bạn có thể sử dụng tổ hợp phím tắt **Ctrl + H** để mở công cụ Replace.


![](https://hddtvn.com/wp-content/uploads/2021/01/9-1.png)


Lúc này, hộp thoại Find and Replace hiện ra. Các bạn chọn mục **More >>**.


![](https://hddtvn.com/wp-content/uploads/2021/01/10-1.png)


Tiếp theo, các bạn chọn mục **Special** => **Graphic**.


![](https://hddtvn.com/wp-content/uploads/2021/01/11.png)


Sau đó, các bạn nhấn vào mục **Replace All**. Lúc này sẽ có thông báo số ảnh sẽ bị xóa đi hiện lên. Như trong trường hợp này là 14 ảnh. Các bạn nhấn **Yes** để thực hiện xóa ảnh. Chỉ cần như vậy là tất cả ảnh trong file Word sẽ bị xóa đi một cách nhanh chóng.


![](https://hddtvn.com/wp-content/uploads/2021/01/13-1.png)


### 2. Xóa ảnh hàng loạt trong Word bằng code VBA


Đầu tiên, các bạn cần chọn thẻ **Developer** trên thanh công cụ. Sau đó các bạn nhấn vào **Visual Basic** tại mục **Code**. Hoặc các bạn có thể sử dụng tổ hợp phím tắt **Alt + F11** để mở cửa sổ VBA.


![](https://hddtvn.com/wp-content/uploads/2021/01/u7cNLHq.png)


Lúc này, cửa sổ Microsoft Visual Basic for Applications hiện ra. Các bạn chọn thẻ **Insert** trên thanh công cụ. Thanh cuộn hiện ra thì các bạn chọn mục **Module**.


![](https://hddtvn.com/wp-content/uploads/2021/01/pzC1xUr.png)


Lúc này, hộp thoại Module hiện ra. Các bạn sao chép đoạn code dưới đây vào đó:


Sub DitchPictures()  

Dim objPic As InlineShape  

For Each objPic In ActiveDocument.InlineShapes  

objPic.Delete  

Next objPic  

End Sub


Sau đó các bạn nhấn vào mục **Run** trên thanh công cụ hoặc sử dụng phím tắt **F5** để chạy mã code trên. Chỉ cần như vậy là tất cả ảnh trong file Word đã được xóa đi một cách nhanh chóng.


![](https://hddtvn.com/wp-content/uploads/2021/01/mLfRrsx.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách xóa ảnh hàng loạt trong Word. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


#### <a href=”http://


moreThông thường, để xóa 1 hình ảnh trong Word thì ta chỉ cần nhấn Delete…

