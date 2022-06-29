package main

import "fmt"

func main() {
	ch := make(chan string, 10)
	var staff = []string{"a", "b", "c", "d", "e", "f", "g", "h", "i", "j"}
	for _, v := range staff {
		ch <- v
	}

	var sms = []int64{1, 2, 3, 4, 5, 6, 7, 8, 9, 10}
	for _, v := range sms {
		s := <-ch
		fmt.Println(s, v)
		ch <- s
	}

}
