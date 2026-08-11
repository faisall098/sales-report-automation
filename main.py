from report_generator import generate_report


def show_menu():
    print()
    print("===== SALES REPORT AUTOMATION =====")
    print("1. Generate Sales Report")
    print("2. Exit")
    print()


def main():

    while True:

        show_menu()

        choice = input("Enter your choice: ")

        if choice == "1":
            generate_report()

        elif choice == "2":
            print("Goodbye!")
            break

        else:
            print("Invalid choice. Please try again.")


if __name__ == "__main__":
    main()