#pragma once
#include "BiosInfo.h"

namespace BIOSExplorer {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// ������ ��� MainForm
	/// </summary>
	public ref class MainForm : public System::Windows::Forms::Form
	{
	public:
		MainForm(void)
		{
			InitializeComponent();
			//
			//TODO: �������� ��� ������������
			//
		}

	protected:
		/// <summary>
		/// ���������� ��� ������������ �������.
		/// </summary>
		~MainForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Label^ label1;
	private: System::Windows::Forms::TextBox^ PrimaryBIOS_tb;
	private: System::Windows::Forms::TextBox^ ReleaseDate_tb;


	private: System::Windows::Forms::Label^ label2;
	private: System::Windows::Forms::TextBox^ SMBIOSPresent_tb;

	private: System::Windows::Forms::Label^ label3;
	private: System::Windows::Forms::TextBox^ Status_tb;

	private: System::Windows::Forms::Label^ label4;
	private: System::Windows::Forms::TextBox^ BIOSVersion_tb;

	private: System::Windows::Forms::Label^ label5;
	private: System::Windows::Forms::RichTextBox^ richBiosCharacteristics;

	private: System::Windows::Forms::Label^ label6;
	private: System::Windows::Forms::Button^ Exit_Btn;
	private: System::Windows::Forms::TextBox^ SMBIOSVersion_tb;



	private: System::Windows::Forms::Label^ label7;
	protected:

	private:
		/// <summary>
		/// ������������ ���������� ������������.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// ��������� ����� ��� ��������� ������������ � �� ��������� 
		/// ���������� ����� ������ � ������� ��������� ����.
		/// </summary>
		void InitializeComponent(void)
		{
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->PrimaryBIOS_tb = (gcnew System::Windows::Forms::TextBox());
			this->ReleaseDate_tb = (gcnew System::Windows::Forms::TextBox());
			this->label2 = (gcnew System::Windows::Forms::Label());
			this->SMBIOSPresent_tb = (gcnew System::Windows::Forms::TextBox());
			this->label3 = (gcnew System::Windows::Forms::Label());
			this->Status_tb = (gcnew System::Windows::Forms::TextBox());
			this->label4 = (gcnew System::Windows::Forms::Label());
			this->BIOSVersion_tb = (gcnew System::Windows::Forms::TextBox());
			this->label5 = (gcnew System::Windows::Forms::Label());
			this->richBiosCharacteristics = (gcnew System::Windows::Forms::RichTextBox());
			this->label6 = (gcnew System::Windows::Forms::Label());
			this->Exit_Btn = (gcnew System::Windows::Forms::Button());
			this->SMBIOSVersion_tb = (gcnew System::Windows::Forms::TextBox());
			this->label7 = (gcnew System::Windows::Forms::Label());
			this->SuspendLayout();
			// 
			// label1
			// 
			this->label1->AutoSize = true;
			this->label1->Location = System::Drawing::Point(12, 68);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(75, 13);
			this->label1->TabIndex = 0;
			this->label1->Text = L"Primary BIOS :";
			// 
			// PrimaryBIOS_tb
			// 
			this->PrimaryBIOS_tb->Location = System::Drawing::Point(128, 65);
			this->PrimaryBIOS_tb->Name = L"PrimaryBIOS_tb";
			this->PrimaryBIOS_tb->ReadOnly = true;
			this->PrimaryBIOS_tb->Size = System::Drawing::Size(105, 20);
			this->PrimaryBIOS_tb->TabIndex = 1;
			// 
			// ReleaseDate_tb
			// 
			this->ReleaseDate_tb->Location = System::Drawing::Point(341, 65);
			this->ReleaseDate_tb->Name = L"ReleaseDate_tb";
			this->ReleaseDate_tb->ReadOnly = true;
			this->ReleaseDate_tb->Size = System::Drawing::Size(131, 20);
			this->ReleaseDate_tb->TabIndex = 3;
			// 
			// label2
			// 
			this->label2->AutoSize = true;
			this->label2->Location = System::Drawing::Point(259, 68);
			this->label2->Name = L"label2";
			this->label2->Size = System::Drawing::Size(78, 13);
			this->label2->TabIndex = 2;
			this->label2->Text = L"Release Date :";
			// 
			// SMBIOSPresent_tb
			// 
			this->SMBIOSPresent_tb->Location = System::Drawing::Point(128, 13);
			this->SMBIOSPresent_tb->Name = L"SMBIOSPresent_tb";
			this->SMBIOSPresent_tb->ReadOnly = true;
			this->SMBIOSPresent_tb->Size = System::Drawing::Size(105, 20);
			this->SMBIOSPresent_tb->TabIndex = 5;
			// 
			// label3
			// 
			this->label3->AutoSize = true;
			this->label3->Location = System::Drawing::Point(3, 16);
			this->label3->Name = L"label3";
			this->label3->Size = System::Drawing::Size(93, 13);
			this->label3->TabIndex = 4;
			this->label3->Text = L"SMBIOS Present :";
			// 
			// Status_tb
			// 
			this->Status_tb->Location = System::Drawing::Point(128, 39);
			this->Status_tb->Name = L"Status_tb";
			this->Status_tb->ReadOnly = true;
			this->Status_tb->Size = System::Drawing::Size(105, 20);
			this->Status_tb->TabIndex = 7;
			// 
			// label4
			// 
			this->label4->AutoSize = true;
			this->label4->Location = System::Drawing::Point(12, 42);
			this->label4->Name = L"label4";
			this->label4->Size = System::Drawing::Size(43, 13);
			this->label4->TabIndex = 6;
			this->label4->Text = L"Status :";
			// 
			// BIOSVersion_tb
			// 
			this->BIOSVersion_tb->Location = System::Drawing::Point(128, 91);
			this->BIOSVersion_tb->Name = L"BIOSVersion_tb";
			this->BIOSVersion_tb->ReadOnly = true;
			this->BIOSVersion_tb->Size = System::Drawing::Size(344, 20);
			this->BIOSVersion_tb->TabIndex = 9;
			// 
			// label5
			// 
			this->label5->AutoSize = true;
			this->label5->Location = System::Drawing::Point(12, 94);
			this->label5->Name = L"label5";
			this->label5->Size = System::Drawing::Size(76, 13);
			this->label5->TabIndex = 8;
			this->label5->Text = L"BIOS Version :";
			// 
			// richBiosCharacteristics
			// 
			this->richBiosCharacteristics->Location = System::Drawing::Point(128, 117);
			this->richBiosCharacteristics->Name = L"richBiosCharacteristics";
			this->richBiosCharacteristics->ReadOnly = true;
			this->richBiosCharacteristics->Size = System::Drawing::Size(344, 215);
			this->richBiosCharacteristics->TabIndex = 10;
			this->richBiosCharacteristics->Text = L"";
			// 
			// label6
			// 
			this->label6->AutoSize = true;
			this->label6->Location = System::Drawing::Point(12, 120);
			this->label6->Name = L"label6";
			this->label6->Size = System::Drawing::Size(110, 13);
			this->label6->TabIndex = 11;
			this->label6->Text = L"BIOS Characteristics :";
			// 
			// Exit_Btn
			// 
			this->Exit_Btn->Location = System::Drawing::Point(15, 309);
			this->Exit_Btn->Name = L"Exit_Btn";
			this->Exit_Btn->Size = System::Drawing::Size(107, 23);
			this->Exit_Btn->TabIndex = 12;
			this->Exit_Btn->Text = L"Exit";
			this->Exit_Btn->UseVisualStyleBackColor = true;
			this->Exit_Btn->Click += gcnew System::EventHandler(this, &MainForm::Exit_Btn_Click);
			// 
			// SMBIOSVersion_tb
			// 
			this->SMBIOSVersion_tb->Location = System::Drawing::Point(341, 39);
			this->SMBIOSVersion_tb->Name = L"SMBIOSVersion_tb";
			this->SMBIOSVersion_tb->ReadOnly = true;
			this->SMBIOSVersion_tb->Size = System::Drawing::Size(131, 20);
			this->SMBIOSVersion_tb->TabIndex = 15;
			// 
			// label7
			// 
			this->label7->AutoSize = true;
			this->label7->Location = System::Drawing::Point(245, 42);
			this->label7->Name = L"label7";
			this->label7->Size = System::Drawing::Size(92, 13);
			this->label7->TabIndex = 14;
			this->label7->Text = L"SMBIOS Version :";
			// 
			// MainForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(484, 339);
			this->Controls->Add(this->SMBIOSVersion_tb);
			this->Controls->Add(this->label7);
			this->Controls->Add(this->Exit_Btn);
			this->Controls->Add(this->label6);
			this->Controls->Add(this->richBiosCharacteristics);
			this->Controls->Add(this->BIOSVersion_tb);
			this->Controls->Add(this->label5);
			this->Controls->Add(this->Status_tb);
			this->Controls->Add(this->label4);
			this->Controls->Add(this->SMBIOSPresent_tb);
			this->Controls->Add(this->label3);
			this->Controls->Add(this->ReleaseDate_tb);
			this->Controls->Add(this->label2);
			this->Controls->Add(this->PrimaryBIOS_tb);
			this->Controls->Add(this->label1);
			this->FormBorderStyle = System::Windows::Forms::FormBorderStyle::FixedSingle;
			this->MaximizeBox = false;
			this->Name = L"MainForm";
			this->Text = L"BIOS Explorer";
			this->Load += gcnew System::EventHandler(this, &MainForm::MainForm_Load);
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion

		BiosInfoOutput bInfOut;

		private: System::Void MainForm_Load(System::Object^ sender, System::EventArgs^ e) {

			if (!bInfOut.ConnectToWMI()) {
				MessageBox::Show(bInfOut.GetResult(), L"�������",
					MessageBoxButtons::OK, MessageBoxIcon::Error);

				// ���������� ��������, �������� �� �������
				Environment::Exit(1);
			}

			MessageBox::Show(bInfOut.GetResult(), L"ϳ���������",
				MessageBoxButtons::OK, MessageBoxIcon::Information);

			this->richBiosCharacteristics->AppendText(bInfOut.GetBiosCharacteristics());
		}

		private: System::Void Exit_Btn_Click(System::Object^ sender, System::EventArgs^ e) {
			// ���������� ��������
			Application::Exit();
		}
};
}
