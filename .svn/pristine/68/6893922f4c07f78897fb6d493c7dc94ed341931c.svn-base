package cl.intelidata.utils;

import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.logging.Logger;

import cl.intelidata.CreateAllExcelAttachment;

public class StringLib {

	private static final Logger	LOGGER	= Logger.getLogger(CreateAllExcelAttachment.class.getName());

	/**
	 * Unifica los valores contenidos dentro de una lista
	 * @param myList
	 * @return
	 */
	public static Set<String> sortList(List<String> myList) {
		Set<String> hashsetList = new HashSet<String>(myList);
		return hashsetList;
	}

	/**
	 * Genera log tipo información
	 * @param msg
	 */
	public static void generateInfo(String msg) {
		LOGGER.info(msg);
		System.out.println(msg);
	}

	/**
	 * Genera log tipo Error
	 * @param alert
	 */
	public static void generateAlert(String alert) {
		LOGGER.severe(alert);
		System.out.println(alert);
	}

	/**
	 * Genera log tipo alerta
	 * @param warning
	 */
	public static void generateWarning(String warning) {
		LOGGER.warning(warning);
		System.out.println(warning);
	}
}
