package main

import (
	"bytes"
	"fmt"
	"log"
	"net/smtp"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

const (
	FILE_NAME       string = "Test2.xlsx"
	SHEET_NAME      string = "Grades"
	SENDER_MAIL     string = ""
	SENDER_PASSWORD string = ""
)

type Score struct {
	Title   string
	Value   string
	Comment string
}

type Student struct {
	ID      string
	Surname string
	Name    string
	Scores  []Score
	Mail    string
}

func (s Student) ToString() string {
	scoresStr := []string{}
	for _, score := range s.Scores {
		scoreStr := fmt.Sprintf("<br><b>%s</b> <br> <b>Score</b>: %s <br> <b>Comment</b>: %s<br>------<br>", score.Title, score.Value, score.Comment)
		scoresStr = append(scoresStr, scoreStr)
	}

	return fmt.Sprintf("ID: <b>%s</b>, Surname: <b>%s</b>, Name: <b>%s</b>, Mail: %s,<br>  -Grades:<br> %s",
		s.ID, s.Surname, s.Name, s.Mail, strings.Join(scoresStr, ""))
}

func read_students_from_xlsx() []Student {
	f, err := excelize.OpenFile(FILE_NAME)
	if err != nil {
		fmt.Println(err)
		return nil
	}

	var students []Student

	rows, err := f.GetRows(SHEET_NAME)
	comments, err := f.GetComments(SHEET_NAME)
	commentMap := make(map[string]string)
	for _, comment := range comments {
		println(comment.Text)
		commentMap[comment.Cell] = comment.Text
	}
	headers := rows[0]
	for i, row := range rows {
		if len((row)) == 0 {
			break
		}
		if i == 0 {
			continue
		}
		student := Student{
			ID:      row[0],
			Surname: row[1],
			Name:    row[2],
			Mail:    row[3],
			Scores:  nil,
		}
		for j, cell := range row {
			if j > 3 {
				cell_id := string(rune(j+65)) + strconv.Itoa(i+1)
				score := Score{
					Title:   headers[j],
					Value:   cell,
					Comment: " ",
				}

				if comment, ok := commentMap[cell_id]; ok {
					score.Comment = comment
				}
				student.Scores = append(student.Scores, score)
			}
		}

		students = append(students, student)
	}
	return students
}

func send_mail(mail, intro, text string) string {
	receivers := []string{mail}
	from := SENDER_MAIL
	password := SENDER_PASSWORD
	smtpHost := "smtp.gmail.com"
	smtpPort := "587"

	var body bytes.Buffer

	mimeHeaders := "MIME-version: 1.0;\nContent-Type: text/html; charset=\"UTF-8\";\n\n"
	body.Write([]byte(fmt.Sprintf("Subject:[No Reply] DSD Grades \n%s\n\n", mimeHeaders) + intro + text))

	fmt.Println("Email is Sending!")

	auth := smtp.PlainAuth("", from, password, smtpHost)
	err := smtp.SendMail(smtpHost+":"+smtpPort, auth, from, receivers, body.Bytes())

	if err != nil {
		log.Fatal(err)
		return "Failed!"
	}
	return ("Email Sent Successfully!")
}

func main() {
	students := read_students_from_xlsx()
	for _, stu := range students {
		text := stu.ToString()
		println(text)
		intro := fmt.Sprintf("Dear <b>%s</b>,<br><br>We hope this message finds you well. We are writing to inform you that your scores have now been recorded and are detailed below."+
			"<br><br>We kindly request that you <b>do not respond directly to this email</b>, even if you wish to appeal any of the scores."+
			"<br><br>Thank you for your understanding and cooperation."+
			"<br>Best Regards<br><br>", stu.Name)
		response := send_mail(stu.Mail, intro, text)
		log.Println(response + "\n")

	}
}
