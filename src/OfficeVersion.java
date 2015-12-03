/** Classe para obter a versão instalada do Microsoft Office em um sistema Windows
 * 
 *  @since IFES-Serra-BSI-TEIE-2015-2
 *  @author Everton Sales
 *  @version 1.02
 *  @see Classe dependente da biblioteca jna (Java Native Access)
 *  @link https://github.com/java-native-access/jna
 */

package br.edu.ifes.bsi.teie;

import com.sun.jna.Platform;
import com.sun.jna.platform.win32.Advapi32Util;
import static com.sun.jna.platform.win32.WinReg.HKEY_CLASSES_ROOT;
import java.io.File;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;

// --------------------------------------------------------------------
// Obtem versão do Office
// --------------------------------------------------------------------
public class OfficeVersion {

	private String version = "";
	
	private int bits = 32;
	
	private int num = 0;
	
	private Boolean ok = false;
	
	private byte[] word;
	
	/** Método para retorno do Nome da versão
     *  @return String - nome da versão do Office */
	public String getVersionName() {
		return version;
	}
	
	/** Método para retorno do Número da versão
     *  @return int - numero da versão do Office */
	public int getVersionNum() {
		return num;
	}
	
	/** Método para retorno da Arquitetura (32 ou 64 bits)
     *  @return int - numero de bits 32 ou 64 */
	public int getVersionBits() {
		return bits;
	}
	
	/** Método para retorno de Execução do construtor bem sucedida
     *  @return Boolean - ok */
	public Boolean isOk(){
		return ok;
	}

	// --------------------------------------------------------------------
	// Construtor
	// --------------------------------------------------------------------
	public OfficeVersion() {
		
		String chave, arquivo;
		
		try
		{
			// Se não está no Windows
			if (!Platform.isWindows()){
				System.out.println( "\nDesculpe-nos, esta versao somente funciona no Windows\n");
				return;
			}
			// Obtem caminho do winword.exe no registro do windows
			if (Advapi32Util.registryKeyExists(HKEY_CLASSES_ROOT, "Applications\\Winword.exe\\shell\\edit\\command"))
			{
				chave = Advapi32Util.registryGetStringValue(HKEY_CLASSES_ROOT, "Applications\\Winword.exe\\shell\\edit\\command", "").toLowerCase();
				arquivo = chave.substring(1, chave.indexOf(".exe")+4);

				// Se o word foi carregado
				if (loadWord(arquivo))
				{
					// Carrega numero da versão
					this.num = loadVersionNum();
				
					// Se Office > 2007
					if (num > 12)
						// Se 64 bits	
						if (is64())
							this.bits = 64;
				
					// converte numero da versão em nome
					this.version = loadVersionName(num);
				
					// Se construtor chegou ao fim sem erros
					this.ok = true;
					
					// Libera memoria
					word = null;
				}
			}
			else
				System.out.println( "\nOcorreu um Erro ao obter o caminho do Office.\n");
				
		}
		catch(Exception e)
		{
			System.out.println( "\nOcorreu um Erro ao obter versao do Office.\n");
		}

	}
	
	// --------------------------------------------------------------------
	// Carrega o arquivo do Word na memória
	// --------------------------------------------------------------------
	private Boolean loadWord(String arquivo)
	{
		try
		{
			File arq = new File(arquivo);
			if (!arq.exists())
			{
				System.out.println( "\nMicrosoft Word nao encontrado em ($arquivo).");
				return false;
			}
			ByteArrayOutputStream out = new ByteArrayOutputStream();  
			FileInputStream in = new FileInputStream(arq);
			// Otimização
			in.skip(arq.length()-13000);
			int b;  
			while((b = in.read())>-1){  
				out.write(b);
			}  
			out.close();  
			in.close();  
			this.word = out.toByteArray();  
			return true;
		}
		catch(Exception e)
		{
			System.out.println( "\nOcorreu um Erro ao carregar o Word\n");
			return false;
		}
		
	}

	// --------------------------------------------------------------------
	// Retorna o nome da versão
	// --------------------------------------------------------------------
	private String loadVersionName(int num) {
		
		switch (num) {
			
			case 16:
				return "2016";
			case 15:
				return "2013";
			case 14:
				return "2010";
			case 12:
				return "2007";
			case 11:
				return "2003";
			case 10:
				return "XP";
			default:
				return "";
		}
		
	}
	
	// --------------------------------------------------------------------
	// Pesquisa dentro do winword.exe o número da versão
	// --------------------------------------------------------------------
	private int loadVersionNum()
	{
		int versionNum = 0;
		for (int x = 0; x < word.length; x++)
		{
			if (word[x]== 80){ // P
				if (word[x+2]==114){ // r
					if (word[x+4]==111){ // o
						if (word[x+6]==100){ // d
							if (word[x+8]==117){ // u
								if (word[x+10]== 99){ // c
									if (word[x+12]==116){ // t
										if (word[x+14]== 86){ // V
											if (word[x+16]==101){ // e
												if (word[x+18]==114){ // r
													if (word[x+20]==115){ // s
														if (word[x+22]==105){ // i
															if (word[x+24]==111){ // o
																if (word[x+26]==110){ // n
																	
																	versionNum = Integer.valueOf(String.format("%c%c", word[x+30], word[x+32]));
																	//System.out.printf("Versao: %c%c", word[x+30], word[x+32]);
																	break;
			}}}}}}}}}}}}}}
		}
		return versionNum;
	}

	// --------------------------------------------------------------------
	// Pesquisa no winword.exe se é 64 bits
	// --------------------------------------------------------------------
	private Boolean is64()
	{
		String arquitetura = "";
		for (int x=0; x < word.length; x++)
		{
			if (word[x]== 65){ // A
				if (word[x+1]==114){ // r
					if (word[x+2]== 99){ // c
						if (word[x+3]==104){ // h
							if (word[x+4]==105){ // i
								if (word[x+5]==116){ // t
									if (word[x+6]==101){ // e
										if (word[x+7]== 99){ // c
											if (word[x+8]==116){ // t
												if (word[x+9]==117){ // u
													if (word[x+10]==114){ // r
														if (word[x+11]==101){ // e
															if (word[x+12]== 61){ // =
																arquitetura = String.format("%c%c%c%c", word[x+14], word[x+15], word[x+16], word[x+17]).toUpperCase();
																break;
		}}}}}}}}}}}}}}
		//System.out.println(arquitetura);
		if (arquitetura.equals("AMD6") || arquitetura.equals("IA64"))
			return true;
		
		return false;
	}

	// --------------------------------------------------------------------
	/** Inicializador para console e exemplo de uso */
	// --------------------------------------------------------------------
	public static void main(String[] args) 
	{
		try
		{
			//System.out.println("\nPrograma iniciado com Sucesso");
			
			OfficeVersion officeVersion = new OfficeVersion();

			if (officeVersion.isOk())
				System.out.printf("\nEsta instalado o Microsoft Office %s - versao: %d - %d bits.\n", officeVersion.getVersionName(), officeVersion.getVersionNum(), officeVersion.getVersionBits());
			else
				System.out.println("\nMicrosoft Office nao Instalado\n");
				

			//System.out.println("\nPrograma terminado com sucesso.\n");	
		}
		catch(Exception e)
		{
			System.out.println("\nPrograma terminado com Erro.\n");
		}
	}
}

